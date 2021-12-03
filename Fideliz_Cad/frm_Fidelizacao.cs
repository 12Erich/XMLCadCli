using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace Fideliz_Cad
{
    public partial class frm_Fidelizacao : Form
    {
        string id = string.Empty;

        public frm_Fidelizacao()
        {
            InitializeComponent();

        }

        private void btnGravar_Click(object sender, EventArgs e)
        {
            SaveClientexml();
            LimparDados();
        }

        private void btnAtualizar_Click(object sender, EventArgs e)
        {
            AtualizaClientexml();
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            BuscarClientexml();
        }

        protected void SaveClientexml()
        {
            try
            {
                //Caminho do arquivo XML
                string sXMLFile = "clientes.xml";

                //Criar o DataSet
                DataSet ds = new DataSet();

                //Preenche o DataSet com o XML
                ds.ReadXml(sXMLFile);

                //Verifica se existe extrutura
                if (ds.Tables.Count == 0)
                {
                    //Cria a extrutura (colunas)
                    DataTable dt = new DataTable("cliente");
                    dt.Columns.Add("id");
                    dt.Columns.Add("Nome");
                    dt.Columns.Add("Email");
                    dt.Columns.Add("Telefone");
                    dt.Columns.Add("Observacao");
                    dt.Columns.Add("Massoterapeuta");
                    dt.Columns.Add("Email");
                    dt.Columns.Add("Avaliacao");
                    dt.Columns.Add("Data_Atualizacao");
                    ds.Tables.Add(dt);
                }
                //Cria uma nova linha
                DataRow dRow = ds.Tables[0].NewRow();

                //Seta os valores
                dRow["id"] = Guid.NewGuid(); //Adiciona um identificador único com ID
                dRow["Nome"] = txNome.Text;
                dRow["Email"] = txtEmail.Text;
                dRow["Telefone"] = txTelefone.Text;
                dRow["Observacao"] = txObservação.Text;
                dRow["Massoterapeuta"] = txMassoterapeuta.Text;
                dRow["Avaliacao"] = txAvaliacao.Text;
                dRow["Data_Atualizacao"] = DateTime.Now;

                //Adiciona a linha no DataSet
                ds.Tables["cliente"].Rows.Add(dRow);

                //Salva no XML
                ds.WriteXml(sXMLFile);

                txBuscaTelefone.Text = txTelefone.Text;
                BuscarClientexml();

                MessageBox.Show("Dados salvos com sucesso!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        protected void AtualizaClientexml()
        {
            try
            {

                if (txNome.Text == "" && txtEmail.Text == "" && txTelefone.Text == "" && txMassoterapeuta.Text == "" && txAvaliacao.Text == "" && txObservação.Text == "")
                {
                    //apagar acidentalmente
                }
                else
                {

                    XDocument xmlDoc = XDocument.Load("clientes.xml");

                    var items = from item in xmlDoc.Elements("clientes").Elements("cliente")
                                where item != null && (item.Element("id").Value == id)
                                select item;

                    foreach (var item in items)
                    {
                        item.Element("Nome").Value = txNome.Text;
                        item.Element("Email").Value = txtEmail.Text;
                        item.Element("Telefone").Value = txTelefone.Text;
                        item.Element("Massoterapeuta").Value = txMassoterapeuta.Text;
                        item.Element("Avaliacao").Value = txAvaliacao.Text;
                        item.Element("Observacao").Value = txObservação.Text;
                        item.Element("Data_Atualizacao").Value = DateTime.Now.ToString();

                        txBuscaTelefone.Text = txTelefone.Text;
                    }

                    xmlDoc.Save("clientes.xml");

                    BuscarClientexml();
                    LimparDados();


                    MessageBox.Show("Dados atualizados com sucesso!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        protected void BuscarClientexml()
        {
            try
            {
                //Caminho do arquivo XML
                string sXMLFile = "clientes.xml";
                //Criar o DataSet
                DataSet ds = new DataSet();
                //Preenche o DataSet com o XML
                ds.ReadXml(sXMLFile);

                if (string.IsNullOrEmpty(txBuscaTelefone.Text))
                {
                    if (ds.Tables["cliente"] != null)
                    {
                        dgClietes.DataSource = ds.Tables["cliente"];

                        dgClietes.Columns[0].Visible = false;
                        dgClietes.Columns[1].ReadOnly = true;
                        dgClietes.Columns[2].ReadOnly = true;
                        dgClietes.Columns[3].ReadOnly = true;
                        dgClietes.Columns[4].ReadOnly = true;
                        dgClietes.Columns[5].ReadOnly = true;
                        dgClietes.Columns[6].ReadOnly = true;
                        dgClietes.Columns[7].ReadOnly = true;

                        dgClietes.AutoResizeColumns();
                        dgClietes.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    }

                }
                else
                {
                    DataTable dtRec = ds.Tables["Cliente"];

                    var dt = from row in dtRec.AsEnumerable()
                             where row.Field<string>("Telefone") == txBuscaTelefone.Text
                             select row;

                    if (dt.Any())
                    {
                        dgClietes.DataSource = dt.CopyToDataTable();

                        id = ((DataTable)dgClietes.DataSource).Rows[0]["id"].ToString();
                        txNome.Text = ((DataTable)dgClietes.DataSource).Rows[0]["Nome"].ToString();
                        txTelefone.Text = ((DataTable)dgClietes.DataSource).Rows[0]["Telefone"].ToString();
                        txtEmail.Text = ((DataTable)dgClietes.DataSource).Rows[0]["Email"].ToString();
                        txMassoterapeuta.Text = ((DataTable)dgClietes.DataSource).Rows[0]["Massoterapeuta"].ToString();
                        txObservação.Text = ((DataTable)dgClietes.DataSource).Rows[0]["Observacao"].ToString();
                        txAvaliacao.Text = ((DataTable)dgClietes.DataSource).Rows[0]["Avaliacao"].ToString();

                    }
                    else
                    {
                        LimparDados();
                        dgClietes.DataSource = null;
                        MessageBox.Show("Não existe este Telefone cadastrado!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        protected void DeleteClientexml()
        {
            try
            {
                //Caminho do arquivo XML
                string sXMLFile = "clientes.xml";

                //Criar o DataSet
                DataSet ds = new DataSet();
                //Preenche o DataSet com o XML
                ds.ReadXml(sXMLFile);

                string strID = "h"; //Request.QueryString["id"];
                if (!string.IsNullOrEmpty(strID))
                {
                    //Selecionar e deletar a linha com o ID da QueryString
                    ds.Tables["cliente"].Select(" id = '" + strID + "'")[0].Delete();
                    //Aplicar as alterações no DataSet
                    ds.AcceptChanges();

                    //Salva no XML
                    ds.WriteXml(sXMLFile);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            DataTable dtRec = new DataTable();
            string strFileOut = string.Empty;

            try
            {
                dtRec = (DataTable)dgClietes.DataSource;

                if (dtRec.Rows.Count > 0)
                {
                    strFileOut = "Relatorio_Clientes_Fidelização_" + +DateTime.Now.Day + "_" + DateTime.Now.Month + "_" + DateTime.Now.Year;

                    GeraCsv_DataGrid(dtRec, null, false);

                }
                else
                {
                    throw new Exception("Nenhum registro a ser exportado!");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Falha ao gerar o arquivo\n" + ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void LimparDados()
        {
            txNome.Text = string.Empty;
            txtEmail.Text = string.Empty;
            txMassoterapeuta.Text = string.Empty;
            txTelefone.Text = string.Empty;
            txObservação.Text = string.Empty;
            txAvaliacao.Text = string.Empty;
            txBuscaTelefone.Text = string.Empty;
        }

        public void CriarDataTable(DataTable dt)
        {
            dt.Columns.Add(CriarColuna("Nome", "System.String"));
            dt.Columns.Add(CriarColuna("Email", "System.String"));
            dt.Columns.Add(CriarColuna("Telefone", "System.String"));
            dt.Columns.Add(CriarColuna("Massoterapeuta", "System.String"));
            dt.Columns.Add(CriarColuna("Avaliacao", "System.String"));
            dt.Columns.Add(CriarColuna("Observacao", "System.String"));
            dt.Columns.Add(CriarColuna("Data_Atualizacao", "System.String"));
        }

        public DataColumn CriarColuna(string NomeColuna, string TipColuna)
        {
            DataColumn Dc = new DataColumn();
            Dc.ColumnName = NomeColuna;
            Dc.DataType = System.Type.GetType(TipColuna);
            return Dc;
        }

        private void GravarRows(DataTable dt)
        {
            string Nome = txNome.Text.Trim();
            string Email = txtEmail.Text.Trim();
            string Telefone = txTelefone.Text.Trim();
            string Avaliacao = txAvaliacao.Text.Trim();
            string Observacao = txObservação.Text.Trim();
            string Massoterapeuta = txMassoterapeuta.Text.Trim();
            string Data_Atualizacao = DateTime.Now.ToString();

            DataRow Dr = dt.NewRow();
            Dr[0] = Nome;
            Dr[1] = Email;
            Dr[2] = Telefone;
            Dr[3] = Avaliacao;
            Dr[4] = Observacao;
            Dr[5] = Massoterapeuta;
            Dr[6] = Data_Atualizacao;
            dt.Rows.Add(Dr);
        }


        public static bool GeraCsv_DataGrid(DataTable dtrGrid,
                                string strFile = null,
                                bool boolAppend = false)
        {

            int intCols;
            int intRows;
            int i;
            int j;
            SaveFileDialog fdPath;
            string strValue;
            System.Data.DataTable dtProcesso = dtrGrid;


            try
            {
                if (strFile == null)
                {
                    fdPath = new SaveFileDialog();
                    {
                        var withBlock = fdPath;
                        withBlock.Filter = "Arquivo CSV|*.csv";
                        withBlock.Title = "Salvar Arquivo CSV";
                        withBlock.ShowDialog();
                    }

                    if (fdPath.FileName != null)
                        strFile = fdPath.FileName;
                    else
                        throw new Exception("Nome de Arquivo não informado !");

                    fdPath.Dispose();
                }

                intCols = dtProcesso.Columns.Count - 1;
                intRows = dtProcesso.Rows.Count - 1;

                using (StreamWriter swFile = new StreamWriter(strFile, boolAppend, Encoding.UTF8))
                {
                    for (i = 0; i <= intCols; i++)
                    {
                        swFile.Write(dtProcesso.Columns[i].ColumnName);
                        if (i < intCols)
                            swFile.Write(";");
                    }

                    swFile.WriteLine();

                    for (i = 0; i <= intRows; i++)
                    {
                        if (dtProcesso.Rows[i][0].ToString() != null)
                        {
                            for (j = 0; j <= intCols; j++)
                            {
                                strValue = dtProcesso.Rows[i][j].ToString().Replace(Environment.NewLine, "").Replace("\n", "").Replace("\r", "");
                                swFile.Write(strValue);
                                if (j < intCols)
                                    swFile.Write(";");
                            }
                            swFile.WriteLine();
                        }
                    }

                    swFile.Close();
                }

                MessageBox.Show("Arquivo gerado com sucesso para: " + strFile, "Exportação", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            finally
            {
                dtProcesso.Clear();
                dtProcesso.Dispose();
                dtProcesso = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}




