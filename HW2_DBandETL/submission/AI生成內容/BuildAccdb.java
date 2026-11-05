import com.healthmarketscience.jackcess.*;

import java.io.File;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

/**
 * Step 4：依 schema.tsv 建立 Access 2007+ (.accdb) 資料庫：
 * 欄位型別 / 大小 / 標題 / 必填 / 驗證規則 / 格式 / 查閱，主鍵、索引、唯一索引、複合唯一索引，
 * 關聯與參考完整性；再依父表 → 子表順序匯入資料，最後重新開檔驗證。
 *
 * 用法：java -cp "lib/*:." BuildAccdb <tsv資料夾> <輸出.accdb>
 */
public class BuildAccdb {
    record Field(String table, String name, String caption, String type, int size, boolean required,
                 String key, boolean index, String rule, String ruleText, String format, String desc,
                 String lookupKind, String rowSource, int columnCount, String columnWidths) {}

    public static void main(String[] args) throws Exception {
        Path dir = Path.of(args[0]);
        File out = new File(args[1]);
        if (out.exists()) out.delete();

        LinkedHashMap<String, List<Field>> tables = new LinkedHashMap<>();
        for (String line : Files.readAllLines(dir.resolve("schema.tsv"), StandardCharsets.UTF_8)) {
            if (line.isBlank()) continue;
            String[] p = line.split("\t", -1);
            Field f = new Field(p[0], p[1], p[2], p[3], Integer.parseInt(p[4]), p[5].equals("1"), p[6],
                                p[7].equals("1"), p[8], p[9], p[10], p[11], p[12], p[13],
                                p[14].isEmpty() ? 0 : Integer.parseInt(p[14]), p[15]);
            tables.computeIfAbsent(f.table, k -> new ArrayList<>()).add(f);
        }
        Map<String, List<List<String>>> uniques = new HashMap<>();
        for (String line : Files.readAllLines(dir.resolve("uniques.tsv"), StandardCharsets.UTF_8)) {
            if (line.isBlank()) continue;
            String[] p = line.split("\t");
            uniques.computeIfAbsent(p[0], k -> new ArrayList<>()).add(Arrays.asList(p).subList(1, p.length));
        }

        try (Database db = new DatabaseBuilder(out).setFileFormat(Database.FileFormat.V2010).create()) {
            db.setAllowAutoNumberInsert(true);   // 自動編號沿用 Excel 產生的 1001 / 2001 / 3001

            // ---- 建立資料表
            for (var e : tables.entrySet()) {
                TableBuilder tb = new TableBuilder(e.getKey());
                List<String> pk = new ArrayList<>();
                for (Field f : e.getValue()) {
                    ColumnBuilder cb = new ColumnBuilder(f.name);
                    switch (f.type) {
                        case "TEXT" -> cb.setType(DataType.TEXT).setLengthInUnits(f.size);
                        case "MEMO" -> cb.setType(DataType.MEMO);
                        case "BYTE" -> cb.setType(DataType.BYTE);
                        case "INT" -> cb.setType(DataType.INT);
                        case "LONG" -> cb.setType(DataType.LONG);
                        case "AUTONUMBER" -> cb.setType(DataType.LONG).setAutoNumber(true);
                        case "MONEY" -> cb.setType(DataType.MONEY);
                        case "DATE", "DATETIME" -> cb.setType(DataType.SHORT_DATE_TIME);
                        default -> throw new IllegalArgumentException(f.type);
                    }
                    cb.putProperty(PropertyMap.CAPTION_PROP, f.caption);
                    if (!f.desc.isEmpty()) cb.putProperty(PropertyMap.DESCRIPTION_PROP, f.desc);
                    if (f.required && !f.type.equals("AUTONUMBER")) cb.putProperty(PropertyMap.REQUIRED_PROP, true);
                    if (f.type.equals("TEXT") || f.type.equals("MEMO"))
                        cb.putProperty(PropertyMap.ALLOW_ZERO_LEN_PROP, false);
                    if (!f.rule.isEmpty()) cb.putProperty(PropertyMap.VALIDATION_RULE_PROP, f.rule);
                    if (!f.ruleText.isEmpty()) cb.putProperty(PropertyMap.VALIDATION_TEXT_PROP, f.ruleText);
                    if (!f.format.isEmpty()) cb.putProperty(PropertyMap.FORMAT_PROP, f.format);
                    if (!f.lookupKind.isEmpty()) {          // 查閱：下拉式方塊
                        cb.putProperty("DisplayControl", DataType.INT, (short) 111);
                        cb.putProperty("RowSourceType", DataType.TEXT, f.lookupKind.equals("values") ? "Value List" : "Table/Query");
                        cb.putProperty("RowSource", DataType.MEMO, f.rowSource);
                        cb.putProperty("BoundColumn", DataType.INT, (short) 1);
                        cb.putProperty("ColumnCount", DataType.INT, (short) f.columnCount);
                        cb.putProperty("ColumnHeads", DataType.BOOLEAN, false);
                        if (!f.columnWidths.isEmpty()) cb.putProperty("ColumnWidths", DataType.TEXT, f.columnWidths);
                        cb.putProperty("ListRows", DataType.INT, (short) 16);
                        cb.putProperty("LimitToList", DataType.BOOLEAN, true);
                    }
                    tb.addColumn(cb);
                    if (f.key.equals("PK")) pk.add(f.name);
                }
                tb.addIndex(new IndexBuilder(IndexBuilder.PRIMARY_KEY_NAME).addColumns(pk.toArray(new String[0])).setPrimaryKey());
                for (Field f : e.getValue()) {
                    if (f.key.equals("UQ")) tb.addIndex(new IndexBuilder(f.name).addColumns(f.name).setUnique());
                    else if (f.index) tb.addIndex(new IndexBuilder(f.name).addColumns(f.name));
                }
                for (List<String> u : uniques.getOrDefault(e.getKey(), List.of()))
                    tb.addIndex(new IndexBuilder("UQ_" + String.join("_", u)).addColumns(u.toArray(new String[0])).setUnique());
                tb.toTable(db);
                System.out.println("建立資料表 " + e.getKey());
            }

            // ---- 關聯（一對多，強制參考完整性 + 串聯更新）
            for (var e : tables.entrySet())
                for (Field f : e.getValue())
                    if (f.key.startsWith("FK:")) {
                        String[] pf = f.key.substring(3).split("\\.");
                        new RelationshipBuilder(pf[0], e.getKey())
                            .addColumns(pf[1], f.name)
                            .setReferentialIntegrity()
                            .setCascadeUpdates()
                            .toRelationship(db);
                        System.out.println("建立關聯 " + pf[0] + "." + pf[1] + " 1 ─ ∞ " + e.getKey() + "." + f.name);
                    }

            // ---- 匯入資料（父表先、子表後）
            for (var e : tables.entrySet()) {
                Table t = db.getTable(e.getKey());
                List<Field> fs = e.getValue();
                int n = 0;
                for (String line : Files.readAllLines(dir.resolve(e.getKey() + ".tsv"), StandardCharsets.UTF_8)) {
                    if (line.isEmpty()) continue;
                    String[] v = line.split("\t", -1);
                    Object[] row = new Object[fs.size()];
                    for (int i = 0; i < fs.size(); i++) row[i] = convert(fs.get(i).type, v[i]);
                    t.addRow(row);
                    n++;
                }
                System.out.println("匯入 " + e.getKey() + "：" + n + " 筆");
            }
        }

        // ---- 重新開啟驗證
        try (Database db = new DatabaseBuilder(out).setReadOnly(true).open()) {
            System.out.println("\n== 驗證 " + out.getName() + " (" + db.getFileFormat() + ")");
            for (String name : db.getTableNames()) {
                Table t = db.getTable(name);
                System.out.printf("  %-18s %3d 筆  索引=%s%n", name, t.getRowCount(),
                                  t.getIndexes().stream().filter(i -> !i.isForeignKey())
                                   .map(i -> i.getName() + (i.isPrimaryKey() ? "(PK)" : i.isUnique() ? "(UQ)" : "")).toList());
            }
            for (Relationship r : db.getRelationships())
                System.out.println("  關聯 " + r.getFromTable().getName() + " → " + r.getToTable().getName()
                                   + "  參考完整性=" + r.hasReferentialIntegrity());
            Column c = db.getTable("StudentTable").getColumn("DeptName");
            System.out.println("  DeptName 屬性：Caption=" + c.getProperties().getValue("Caption")
                               + " RowSource=" + c.getProperties().getValue("RowSource"));
            System.out.println("  第一筆報名：" + db.getTable("RegistrationTable").iterator().next());
        }
    }

    static Object convert(String type, String s) {
        if (s.isEmpty()) return null;
        return switch (type) {
            case "TEXT", "MEMO" -> s;
            case "BYTE" -> Byte.parseByte(s);
            case "INT" -> Short.parseShort(s);
            case "LONG", "AUTONUMBER" -> Integer.parseInt(s);
            case "MONEY" -> new BigDecimal(s);
            case "DATE" -> LocalDate.parse(s.substring(0, 10)).atStartOfDay();
            case "DATETIME" -> LocalDateTime.parse(s);
            default -> throw new IllegalArgumentException(type);
        };
    }
}
