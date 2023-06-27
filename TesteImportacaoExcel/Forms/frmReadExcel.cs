using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ExcelDataReader;
using TesteImportacaoExcel.Tabelas;
using Oracle.ManagedDataAccess.Client;

namespace TesteImportacaoExcel
{
    public partial class frmReadExcel : Form
    {
        #region Atributos Privados
        testeImportacao testeImportacao = new testeImportacao();
        #endregion

        #region Construtor
        public frmReadExcel()
        {
            InitializeComponent();
        }
        #endregion

        #region Trata DataRow
        private void TrataDataRow(in string descProd, in string descItem, out string descProdResult, out string descItemResult, out string codProd, out string codItem)
        {
            int indicePrimeiroHifen = 0;

            // Trata descProd
            indicePrimeiroHifen = descProd.IndexOf('-');
            if (indicePrimeiroHifen >= 0)
            {
                codProd = descProd.Substring(0, indicePrimeiroHifen).Trim();
                descProdResult = descProd.Substring(indicePrimeiroHifen + 1).Trim();
            }
            else
            {
                codProd = string.Empty;
                descProdResult = descProd;
            }

            // Trata descItem
            indicePrimeiroHifen = descItem.IndexOf('-');
            if (indicePrimeiroHifen >= 0)
            {
                codItem = descItem.Substring(0, indicePrimeiroHifen).Trim();
                descItemResult = descItem.Substring(indicePrimeiroHifen + 1).Trim();
            }
            else
            {
                codItem = string.Empty;
                descItemResult = descItem;
            }
        }
        #endregion

        #region Importa Excel
        private bool ImportaExcel(DataTable dt)
        {
            bool retorno = true;
            string descProdDefault = "";
            string msgErro = "";
            int numLinha = 1;

            // Criando a Conexão com o bd
            OracleConnection con;
            con = new OracleConnection(Properties.Settings.Default.connStr);

            // Abre a conexao com o bd
            con.Open();

            if (dt.Rows.Count > 0)
            {
                // Obtém a primeira linha
                DataRow firstRow = dt.Rows[0];

                foreach (DataColumn column in dt.Columns)
                {
                    descProdDefault = firstRow[0].ToString();
                }

                foreach (DataRow row in dt.Rows)
                {
                    testeImportacao.DescProd = descProdDefault;
                    testeImportacao.Ordem = Convert.ToInt32(row[1]);
                    testeImportacao.DescItem = row[2].ToString();
                    testeImportacao.TipoItem = row[3].ToString();
                    testeImportacao.UnidadeItem = row[4].ToString();
                    testeImportacao.QtdeItem = Convert.ToDecimal(row[5]);

                    if(row.ItemArray.Length == 7)
                    {
                        testeImportacao.CustoItem = Convert.ToDecimal(row[6]);
                    }
                    else
                        testeImportacao.CustoItem = 0;

                    TrataDataRow(testeImportacao.DescProd, testeImportacao.DescItem, out string descProdResult, out string descItemResult, out string codProd, out string codItem);

                    testeImportacao.DescProd = descProdResult;
                    testeImportacao.DescItem = descItemResult;
                    testeImportacao.CodProd = codProd;
                    testeImportacao.CodItem = codItem;

                    testeImportacao.DefineConexao(con);

                    if (!testeImportacao.Existe())
                    {
                        if (!testeImportacao.Insert(out msgErro))
                        {
                            Logger.Log($"Erro ao inserir a linha {numLinha}, Erro: {msgErro}");
                            // Erro ao inserir dados
                        }
                        else
                        {
                            Logger.Log($"Linha {numLinha} inserida com sucesso!");
                            // Dados inseridos com sucesso
                        }
                    }
                    else
                    {
                        Logger.Log($"Erro ao inserir a linha {numLinha}, os dados já existem");
                        // Os dados já existem
                    }

                    numLinha++;
                }
            }
            else
            {
                return false;
            }
            return retorno;
        }
        #endregion

        #region Botão Open Folder
        private void btnOpenFile_Click(object sender, EventArgs e)
             {
            // Verifica se já existe um arquivo log
            if (File.Exists(@"C:\Users\Hino\Desktop\log.txt"))
            {
                // Remove o arquivo existente
                File.Delete(@"C:\Users\Hino\Desktop\log.txt");
            }

            // Abre o diálogo de seleção da pasta que contém os arquivos Excel
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    string folderPath = folderBrowserDialog.SelectedPath;
                    string[] excelFiles = Directory.GetFiles(folderPath, "*.xls*");
                    bool ocorreramErros = false;

                    foreach (string excelFile in excelFiles)
                    {
                        // Abre o arquivo Excel em modo de leitura
                        using (var stream = File.Open(excelFile, FileMode.Open, FileAccess.Read))
                        {
                            // Cria um leitor de Excel usando a biblioteca ExcelDataReader
                            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                // Lê o Excel e passa para o DataSet result
                                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                                {
                                    // Configura a leitura das tabelas do Excel considerando a primeira linha como cabeçalho
                                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                    {
                                        UseHeaderRow = true,
                                        ReadHeaderRow = rowReader =>
                                        {
                                            rowReader.Read();
                                        }
                                    }
                                });

                                // Armazena a tabela do arquivo Excel no DataTable
                                DataTable dt = result.Tables[0];

                                Logger.Log($"Arquivo {excelFile}");

                                // Importar para o banco
                                if (!ImportaExcel(dt))
                                {
                                    ocorreramErros = true;
                                }
                            }
                        }
                    }

                    if (ocorreramErros)
                    {
                        MessageBox.Show("Importação não concluída, arquivo log gerado");
                    }
                    else
                    {
                        MessageBox.Show("Importação concluída, arquivo log gerado");
                    }
                }
            }
        }
        #endregion

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            List<Item> itens = new List<Item>()
            {
                new Item() { Ordem = 1, CodProduto = "A98", Tipo = "SEMI" },
                new Item() { Ordem = 2, CodProduto = "TA98", Tipo = "SEMI" },
                new Item() { Ordem = 3, CodProduto = "MA98", Tipo = "SEMI" },
                new Item() { Ordem = 4, CodProduto = "11012907", Tipo = "MP" },
                new Item() { Ordem = 5, CodProduto = "999", Tipo = "MOT" },
                new Item() { Ordem = 6, CodProduto = "ABC", Tipo = "SEMI" },
                new Item() { Ordem = 7, CodProduto = "DEF", Tipo = "SEMI" },
            };

            IdentificarPai identificarPai = new IdentificarPai(itens);
            int minOrdem = itens.Min(r => r.Ordem);
            Item primeiroItem = itens.Where(r => r.Ordem == minOrdem).First();
            identificarPai.Processar(primeiroItem.CodProduto, primeiroItem);

            foreach (var item in identificarPai._itens)
            {
                Console.WriteLine($"Pai - {item.CodPai} | Produto {item.CodProduto} ");
            }
        }
    }
}
