package importardatosexcel;

import java.sql.SQLException;

public class ImportarDatosExcel {

    public static void main(String[] args) throws SQLException, ClassNotFoundException {
//        Cuentas de socio prueba
//        CuentasSocios. importarDatos("C:\\Users\\Fabian\\Documents\\Sigma\\UnionProgreso\\CUENTAS_SOCIOS.xlsx", 1, 8);
//        Cunetas de socios
        Cuentas. importarDatos("C:\\Users\\Fabian\\Documents\\Sigma\\UnionProgreso\\CUENTAS_SOCIOS.xlsx", 20, 8);
//        Creditos Emergentes Cuotas Mensulaes
//        CreditoEmergente.importarDatos("C:\\Users\\Fabian\\Documents\\Sigma\\UnionProgreso\\Cuotas_Créditos_Emergentes.xlsx", 16, 12, 4);
//        EstadoCuenta.importarDatos("C:\\Users\\Fabian\\Documents\\Sigma\\JardinAzuayo\\Estado_Cuenta.xlsx", 1, 14, 4);
    }
}
