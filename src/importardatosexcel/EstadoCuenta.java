package importardatosexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EstadoCuenta {

    private static Connection cn = null;

    public static ArrayList<String[]> importarDatos(String fileName, int numSheet, int numColums, int numRow) throws SQLException, ClassNotFoundException {

        ArrayList<String[]> data = new ArrayList<>();

        for (int hoja = 0; hoja < numSheet; hoja++) {

            try {
                FileInputStream file = new FileInputStream(new File(fileName));
                XSSFWorkbook workbook = new XSSFWorkbook(file);
                XSSFSheet sheet = workbook.getSheetAt(hoja);
                Iterator<Row> rowIterator = sheet.iterator();
                int numFila = 0;
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();
                    String[] fila = new String[numColums];
                    if (numFila > numRow - 2 && row.getCell(0) != null) {
                        int numColumna = 0;
                        while (cellIterator.hasNext()) {
                            Cell cell = cellIterator.next();
                            if (numColumna < numColums) {
                                switch (cell.getCellType()) {
                                    case Cell.CELL_TYPE_NUMERIC:
                                        if (DateUtil.isCellDateFormatted(cell)) {
                                            fila[numColumna] = new SimpleDateFormat("dd/MM/yyyy").format(cell.getDateCellValue()) + "";
                                        } else {
                                            fila[numColumna] = cell.getNumericCellValue() + "";
                                        }
                                        break;
                                    case Cell.CELL_TYPE_STRING:
                                        fila[numColumna] = cell.getStringCellValue() + "";
                                        break;
                                    case Cell.CELL_TYPE_BLANK:
                                        fila[numColumna] = "";
                                        break;
                                    case Cell.CELL_TYPE_FORMULA:
                                        fila[numColumna] = "" + getDecimal(2, cell.getNumericCellValue());
                                        break;
                                    case Cell.CELL_TYPE_ERROR:
                                        fila[numColumna] = "ERR";
                                        break;
                                    default:
                                        fila[numColumna] = cell.toString();
                                        System.out.println("Ingreso " + cell.toString());
                                        break;
                                }
//                                fila[numColumna] = "" + cell.toString();
                            }
                            numColumna++;
                        }
                        if (fila[0] != null) {
                            if (!fila[0].equals("")) {
                                data.add(fila);
                            }
                        }
                    }
                    numFila++;
                }
                conexio();
                for (int i = 0; i < data.size(); i++) {
                    if (!data.get(i)[0].equalsIgnoreCase("FECHA")) {
//                        for (int j = 0; j < numColums; j++) {
//                            System.out.print(j+" "+data.get(i)[j] + "\t");
//                        }
                        System.out.println(data.get(i)[0]+" "+data.get(i)[2]+" "+data.get(i)[4]+" "+data.get(i)[6]+" "+ data.get(i)[7]+" "+ data.get(i)[8]+" "+ data.get(i)[9]+" "+ data.get(i)[11]+" "+ data.get(i)[12]+" "+ data.get(i)[13]);
//                        guardarDatos(data.get(i)[0], data.get(i)[2], data.get(i)[4], data.get(i)[6], data.get(i)[7], data.get(i)[8], data.get(i)[9], data.get(i)[11], data.get(i)[12], data.get(i)[13]);
                    }
//                    System.out.println("");
                }
                cn.close();
                data.clear();
            } catch (IOException e) {
                System.out.println("ERROR\n" + e);
            }
        }
        System.out.println("Excel importado correctamente\n");
        return data;
    }

    public static double getDecimal(int numeroDecimales, double decimal) {
        decimal = decimal * (java.lang.Math.pow(10, numeroDecimales));
        decimal = java.lang.Math.round(decimal);
        decimal = decimal / java.lang.Math.pow(10, numeroDecimales);

        return decimal;
    }

    public static void guardarDatos(String p1, String p2, String p3, String p4, String p5, String p6, String p7, String p8, String p9, String p10) throws ClassNotFoundException {
        if (cn != null) {
            try {
                CallableStatement cst = cn.prepareCall("{call pcrear_estado_cuenta (?,?,?,?,?,?,?,?,?,?)}");
                cst.setString(1, p1);
                cst.setString(2, p2);
                cst.setString(3, p3);
                cst.setString(4, p4);
                cst.setString(5, p5);
                cst.setString(6, p6);
                cst.setString(7, p7);
                cst.setString(8, p8);
                cst.setString(9, p9);
                cst.setString(10, p10);
                cst.execute();
                cst.close();
            } catch (SQLException e) {
                System.err.println("ERROR en el procedimiento de la cabecera\n" + e);
            }
        } else {
            System.out.println("No hay conexion");
        }

    }

    public static void conexio() {
        try {
            Class.forName("oracle.jdbc.OracleDriver");
            cn = DriverManager.getConnection("jdbc:oracle:thin:@localhost:1521:XE", "sigma", "sigma");
        } catch (ClassNotFoundException | SQLException e) {
            System.out.println("ERROR en la conexion a la BD\n" + e);
        }
    }
}
