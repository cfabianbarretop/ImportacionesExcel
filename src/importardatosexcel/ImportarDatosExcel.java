package importardatosexcel;

import java.sql.SQLException;

public class ImportarDatosExcel {

    public static void main(String[] args) throws SQLException, ClassNotFoundException {
//        Cuentas. importarDatos("C:\\Users\\Fabian\\Documents\\Sigma\\UnionProgreso\\cuentasP.xlsx", 3, 8);
//        CreditoEmergente.importarDatos("C:\\Users\\Fabian\\Documents\\Sigma\\UnionProgreso\\Cuotas_Créditos_Emergentes.xlsx", 3, 12, 4);
        EstadoCuenta.importarDatos("C:\\Users\\Fabian\\Documents\\Sigma\\JardinAzuayo\\Estado_Cuenta.xlsx", 1, 14, 4);
    }
}
