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

public class Cuentas {

    private static Connection cn = null;

    public static ArrayList<String[]> importarDatos(String fileName, int numSheet, int numRow) throws SQLException, ClassNotFoundException {
        ArrayList<String[]> data = new ArrayList<>();
        for (int hoja = 0; hoja < numSheet; hoja++) {
            try {
                FileInputStream file = new FileInputStream(new File(fileName));
                XSSFWorkbook workbook = new XSSFWorkbook(file);
                XSSFSheet sheet = workbook.getSheetAt(hoja);
                Iterator<Row> rowIterator = sheet.iterator();

                int fila = 0, columna = 0, tablas = 0, i = 0;

                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();
                    i = 0;
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        i++;
                    }
                    if (fila == 12) {
                        columna = i;
                    }
                    fila++;
                }
                tablas = columna / 8;
//            tablas+=1;//VALOR AUMENTA EN SER NECESARIO
                int celda = 0;

                for (int tab = 0; tab < tablas; tab++) {
                    Iterator<Row> rowIter = sheet.iterator();
                    System.out.println("------------------Tabla " + (tab + 1) + "--------------------");
                    for (int row = 0; row < fila; row++) {
                        String[] Fila = new String[8];
                        int posc = 0;
                        Row fil = rowIter.next();
                        if (row > 11 || row == 2) {
                            for (int colum = celda; colum < celda + 8; colum++) {
                                if (fil.getCell(colum) != null) {
                                    switch (fil.getCell(colum).getCellType()) {
                                        case Cell.CELL_TYPE_NUMERIC:
                                            if (DateUtil.isCellDateFormatted(fil.getCell(colum))) {
                                                Fila[posc] = new SimpleDateFormat("dd/MM/yyyy").format(fil.getCell(colum).getDateCellValue()) + "";
                                            } else {
                                                Fila[posc] = fil.getCell(colum).getNumericCellValue() + "";
                                            }
                                            break;
                                        case Cell.CELL_TYPE_STRING:
                                            Fila[posc] = fil.getCell(colum).getStringCellValue() + "";
                                            break;
                                        case Cell.CELL_TYPE_BLANK:
                                            Fila[posc] = "";
                                            break;
                                        case Cell.CELL_TYPE_FORMULA:
                                            Fila[posc] = getDecimal(2, fil.getCell(colum).getNumericCellValue()) + "";
                                            break;
                                        case Cell.CELL_TYPE_ERROR:
                                            Fila[posc] = fil.getCell(colum) + "&";
                                            break;
                                        default:
                                            Fila[posc] = fil.getCell(colum) + "&";
                                            break;
                                    }
                                }
                                posc++;
                            }
                            if (Fila[0] != null) {
                                if (!Fila[0].equals("")) {
                                    data.add(Fila);
                                }
                            }
                        }
                    }
                    String cuenta = "";
                    String nombre = "";
                    conecxio();
                    for (int x = 0; x < data.size(); x++) {
                        if (x == 0) {
                            cuenta = data.get(x)[1];
                            nombre = data.get(x)[5];
                            System.out.println("Cuenta: " + cuenta);
                        } else {
//                            guardarDatos(cuenta, nombre, data.get(x)[0], data.get(x)[1], data.get(x)[2], data.get(x)[3], data.get(x)[4], data.get(x)[5], data.get(x)[6]);
                            System.out.println("| " + cuenta + " | " + nombre + " | " + data.get(x)[0] + " | " + data.get(x)[1] + " | " + data.get(x)[2] + " | " + data.get(x)[3] + " | " + data.get(x)[4] + " | " + data.get(x)[5] + " | " + data.get(x)[6] + " |");
                        }
                    }
                    System.out.println("");
                    cn.close();
                    data.clear();
                    celda += 10;
                }
            } catch (IOException e) {
                System.err.println("ERROR en la importacion de datos\n" + e);
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

    public static void guardarDatos(String cta, String nom, String fech, String trans, String val, String inter, String sal_sc, String sal_usd, String dia
    ) throws ClassNotFoundException {
        try {
            if (cn != null) {
                CallableStatement cst = cn.prepareCall("{call pcrear_cta (?,?,?,?,?,?,?,?,?)}");
                cst.setString(1, cta);
                cst.setString(2, nom);
                cst.setString(3, fech);
                cst.setString(4, trans);
                cst.setString(5, val);
                cst.setString(6, inter);
                cst.setString(7, sal_sc);
                cst.setString(8, sal_usd);
                cst.setString(9, dia);
                cst.execute();
                cst.close();
            } else {
                System.out.println("No hay conexion");
            }
        } catch (SQLException e) {
            System.err.println("ERROR en el procedimiento alnmacenado\n" + e);
        }
    }

    public static void conecxio() {
        try {
            Class.forName("oracle.jdbc.OracleDriver");
            cn = DriverManager.getConnection("jdbc:oracle:thin:@localhost:1521:XE", "sigma", "sigma");
        } catch (Exception e) {
            System.out.println("ERRROR en la conxion de la base de datos\n" + e);
        }
    }

}
