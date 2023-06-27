using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TesteImportacaoExcel.Tabelas
{
    public class testeImportacao
    {
        #region Propriedades privadas
        private String _connStr;
        #endregion

        #region Atributos Públicos
        public string DescProd { get; set; }
        public string CodProd { get; set; }
        public int? Ordem { get; set; }
        public string DescItem { get; set; }
        public string CodItem { get; set; }
        public string TipoItem { get; set; }
        public string UnidadeItem { get; set; }
        public decimal? QtdeItem { get; set; }
        public decimal? CustoItem { get; set; }
        #endregion

        #region Define a conexão
        private OracleConnection _OrclConn;
        public void DefineConexao(OracleConnection pOrclConn)
        {
            this._OrclConn = pOrclConn;
        }
        #endregion

        #region Método Construtor
        public testeImportacao()
        {
        }
        #endregion

        #region Insert
        public Boolean Insert(out String mensErro)
        {
            bool retorno = true;
            mensErro = "";

            // Define a instrução SQL
            string strSql = @"INSERT INTO TESTEIMPORTACAO
                                (DESCPROD, CODPROD, ORDEM, DESCITEM, 
                                CODITEM, TIPOITEM, UNIDADEITEM, QTDEITEM, 
                                CUSTOITEM)
                              VALUES 
                                (:pDESCPROD, :pCODPROD, :pORDEM, :pDESCITEM, 
                                :pCODITEM, :pTIPOITEM, :pUNIDADEITEM, :pQTDEITEM, 
                                :pCUSTOITEM)";

            // Cria a conexão com o banco de dados
            OracleConnection con;
            if (_OrclConn == null)
            {
                con = new OracleConnection(_connStr);
                // Abre a conexao
                con.Open();
            }
            else
            {
                con = _OrclConn;
            }

            // Cria o objeto command para executar a instruçao sql
            using (OracleCommand cmd = new OracleCommand())
            {
                // Define o tipo do comando
                cmd.CommandType = CommandType.Text;

                // Define a conexão
                cmd.Connection = con;

                // Define o comando
                cmd.CommandText = strSql;

                // Parâmetros
                cmd.Parameters.Add("pDESCPROD", OracleDbType.Varchar2).Value = DescProd;
                cmd.Parameters.Add("pCODPROD", OracleDbType.Varchar2).Value = CodProd;
                cmd.Parameters.Add("pORDEM", OracleDbType.Int32).Value = Ordem;
                cmd.Parameters.Add("pDESCITEM", OracleDbType.Varchar2).Value = DescItem;
                cmd.Parameters.Add("pCODITEM", OracleDbType.Varchar2).Value = CodItem;
                cmd.Parameters.Add("pTIPOITEM", OracleDbType.Varchar2).Value = TipoItem;
                cmd.Parameters.Add("pUNIDADEITEM", OracleDbType.Varchar2).Value = UnidadeItem;
                cmd.Parameters.Add("pQTDEITEM", OracleDbType.Decimal).Value = QtdeItem;
                cmd.Parameters.Add("pCUSTOITEM", OracleDbType.Decimal).Value = CustoItem;

                // Executando o comando
                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    mensErro = $"Problema ao incluir o registro. Motivo: {e.Message}";
                    retorno = false;
                }
            }

            // Fecha a conexão
            if (_OrclConn == null)
            {
                OracleConnection.ClearPool(con);
                con.Close();
                con.Dispose();
            }

            return retorno;
        }
        #endregion

        #region Existe
        public Boolean Existe()
        {
            Boolean retorno = true;
            String strSql = @"SELECT COUNT(*) TOTREG
                                FROM TESTEIMPORTACAO
                               WHERE CODPROD = :pCODPROD
                                 AND ORDEM = :pORDEM
                                 AND CODITEM = :pCODITEM";

            // Cria a conexão com o banco de dados
            OracleConnection con;
            if (_OrclConn == null)
            {
                con = new OracleConnection(_connStr);
                // Abre a conexao
                con.Open();
            }
            else
            {
                con = _OrclConn;
            }

            // Cria o objeto command para executar a instruçao sql
            using (OracleCommand cmd = new OracleCommand())
            {
                // Define o tipo do comando
                cmd.CommandType = CommandType.Text;

                // Define a conexão
                cmd.Connection = con;

                // Define o comando
                cmd.CommandText = strSql;

                // Parâmetros
                cmd.Parameters.Add("pCODPROD", OracleDbType.Varchar2).Value = CodProd;
                cmd.Parameters.Add("pORDEM", OracleDbType.Int32).Value = Ordem;
                cmd.Parameters.Add("pCODITEM", OracleDbType.Varchar2).Value = CodItem;

                // Executando o comando
                retorno = (int.Parse(cmd.ExecuteScalar().ToString()) > 0);
            }

            // Fecha a conexão
            if (_OrclConn == null)
            {
                OracleConnection.ClearPool(con);
                con.Close();
                con.Dispose();
            }

            return retorno;
        }
        #endregion
    }
}
