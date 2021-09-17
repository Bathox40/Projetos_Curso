using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Model_VO;
using Camada_BLL;
using Microsoft.VisualBasic.Compatibility;
using Excel = Microsoft.Office.Interop.Excel;
using Email = Microsoft.Office.Interop.Outlook;
using Microsoft.VisualBasic;

namespace AltomacaoComSqlServer
{
    public partial class frmAltomacaoFull : Form
    {
        Cliente_BLL objCliente_BLL;
        Cliente_VO objCliente_VO;

        int intValorSalvo;
        string strValorAntigo;
        bool bolAddBd;
        //--------------------------------------------

        Pedidos_Interior_BLL objPedidos_Interior_BLL;
        Pedidos_Interior_VO objPedidos_Interior_VO;

        int intID_Salvo_PI;
        int intCliente_ID_Salvo_PI;
        string strNome_Salvo_PI;
        bool bolAddBd_PI;
        //----------------------------------------------

        Pedidos_Exterior_BLL objPedidos_Exterior_BLL;
        Pedidos_Exterior_VO objPedidos_Exterior_VO;

        int intID_Salvo_PE;
        int intCliente_ID_Salvo_PE;
        string strNome_Salvo_PE;
        bool bolAddBd_PE;
        //--------------------------------------------

        Excel.Application objApplication;
        Excel.Workbook objWorkbook;
        Excel.Worksheet objWorksheet;
        Excel.Range objCabecalho;
        Excel.Range objExDados;
        //----------------------------------------------

        Email.Application objEmailApp;
        Email.MailItem objEmailMsn;
        Email.OlAttachmentType objAttchment;

        string[] objAnexoArq = new String[0];
        long objAnexoPosition;
        string objDisplayName;

        //--------------------------------------------
        string strPesquisarIdClientes;
        string strQuantidade_de_PI_Cliente;

        public frmAltomacaoFull()
        {
            InitializeComponent();
        }

        // PI = Pedidos Iterior
        // PE = Pedidos Exterior

        #region Cliente
        private void btnEstrutudEsc_Click(object sender, EventArgs e)
        {
            switch (MessageBox.Show("Escolha Uma Opção", "Estrutura de Escolha", MessageBoxButtons.YesNoCancel))
            {
                case DialogResult.Yes:
                    MessageBox.Show("Você Escolheu Sim!");
                    break;
                case DialogResult.No:
                    MessageBox.Show("Você Escolheu Não!");
                    break;
                case DialogResult.Cancel:
                    MessageBox.Show("Você Escolheu Cancelar!");
                    break;
                default:
                    MessageBox.Show("Erro! Escolha Sim, Não ou Cancelar");
                    break;

            }
        }

        private void btnImporTxt_Click(object sender, EventArgs e)
        {
            ImportarTXT();
        }
        public void ImportarTXT()
        {
            try
            {
                lstbxClientes.Items.Clear();
                objCliente_BLL = new Cliente_BLL();
                lstbxClientes.Items.AddRange(objCliente_BLL.ImportarTXT().ToArray());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Presta Atenção !!!  :" + ex.Message);
            }
        }

        private void btnImporBdc_Click(object sender, EventArgs e)
        {
            ImportarBDC();
        }
        public void ImportarBDC()
        {
            try
            {
                lstbxClientes.Items.Clear();
                objCliente_BLL = new Cliente_BLL();
                lstbxClientes.Items.AddRange(objCliente_BLL.ImportarBDC().ToArray());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Presta Atenção !!!  :" + ex.Message);
            }
        }

        private void btnImporBdd_Click(object sender, EventArgs e)
        {
            ImportarBDD();
        }
        public void ImportarBDD()
        {
            try
            {
                lstbxClientes.Items.Clear();
                objCliente_BLL = new Cliente_BLL();
                lstbxClientes.Items.AddRange(objCliente_BLL.ImportarBDD().ToArray());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Presta Atenção !!!  :" + ex.Message);
            }
        }

        private void btnConsBd_Click(object sender, EventArgs e)
        {
            ConsultarBD();
        }

        public void ConsultarBD(int? intID = null, string strNome = null)
        {
            try
            {
                objCliente_BLL = new Cliente_BLL();
                objCliente_VO = new Cliente_VO();

                objCliente_VO.ID = Convert.ToInt32(intID == null ? 0 : intID);
                objCliente_VO.Nome = strNome;

                bndsrcClientes.DataSource = objCliente_BLL.ConsultarBD(objCliente_VO);
                dtgdvwClientes.DataSource = bndsrcClientes;

                cmbbxPI.DataSource = null;
                cmbbxPI.Items.Clear();
                cmbbxPI.DisplayMember = "Nome";
                cmbbxPI.ValueMember = "ID";

                cmbbxPI.DataSource = bndsrcClientes.DataSource;
                cmbbxPI.SelectedIndex = Convert.ToInt32(intID > 0 ? intID - 1 : 0);

                cmbbxPE.DataSource = null;
                cmbbxPE.Items.Clear();
                cmbbxPE.DisplayMember = "Nome";
                cmbbxPE.ValueMember = "ID";

                cmbbxPE.DataSource = bndsrcClientes.DataSource;
                cmbbxPE.SelectedIndex = Convert.ToInt32(intID > 0 ? intID - 1 : 0);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Presta Atenção !!!  :" + ex.Message);
            }
        }

        private void btnIncBd_Click(object sender, EventArgs e)
        {
         IncluirBD(dtgdvwClientes.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(),
                      dtgdvwClientes.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                      Convert.ToInt32(dtgdvwClientes.CurrentRow.Cells["Ativos"].EditedFormattedValue.ToString()));
            ConsultarBD();
        }
        public void IncluirBD(string strNome, string strDescricao, int intAtivo)
        {
            try
            {
                objCliente_BLL = new Cliente_BLL();
                objCliente_VO = new Cliente_VO();

                objCliente_VO.Nome = strNome;
                objCliente_VO.Descricao = strDescricao;
                objCliente_VO.Ativos = intAtivo;

                if (objCliente_BLL.IncluirBD(objCliente_VO))
                {
                    MessageBox.Show("Registro Incluso!");
                }
                else
                {
                    MessageBox.Show("Registro Não Incluso!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Presta Atenção !!!  :" + ex.Message);
            }
        }

        private void btnExclBd_Click(object sender, EventArgs e)
        {
            ExcluirBD(intValorSalvo);
            ConsultarBD();
        }
        public void ExcluirBD(int intID)
        {
            try
            {
                objCliente_BLL = new Cliente_BLL();
                objCliente_VO = new Cliente_VO();

                objCliente_VO.ID = intID;


                if (objCliente_BLL.ExcluirBD(objCliente_VO))
                {
                    MessageBox.Show("Registro Excluido!");
                }
                else
                {
                    MessageBox.Show("Registro Não Excluido!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Presta Atenção !!!  :" + ex.Message);
            }
        }

        private void btnAltBd_Click(object sender, EventArgs e)
        {
            AlterarBD(intValorSalvo, dtgdvwClientes.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(),
                   dtgdvwClientes.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                   Convert.ToInt32(dtgdvwClientes.CurrentRow.Cells["Ativos"].EditedFormattedValue.ToString()));
            ConsultarBD();

        }
        public void AlterarBD(int intID, string strNome, string strDescricao, int intAtivo)
        {
            try
            {
                objCliente_BLL = new Cliente_BLL();
                objCliente_VO = new Cliente_VO();

                objCliente_VO.ID = intID;
                objCliente_VO.Nome = strNome;
                objCliente_VO.Descricao = strDescricao;
                objCliente_VO.Ativos = intAtivo;

                if (objCliente_BLL.AlterarBD(objCliente_VO))
                {
                    MessageBox.Show("Registro Alterado!");
                }
                else
                {
                    MessageBox.Show("Registro Não Alterado!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Presta Atenção !!!  :" + ex.Message);
            }
    }

        private void bndnavbtnAddCL_Click(object sender, EventArgs e)
        {
            bolAddBd = true;
        }

        private void bndnavbtnExcCl_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Excluir" + strValorAntigo + "? ", "Excluindo Registro", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                ExcluirBD(intValorSalvo);
                ConsultarBD();
            }
        }

        private void bndnavbtnConfirmarCL_Click(object sender, EventArgs e)
        {
            if (bolAddBd)
            {
                if (MessageBox.Show("Deseja Incluir" + dtgdvwClientes.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString() + "? ", "Incluindo Registro", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    IncluirBD(dtgdvwClientes.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(),
                      dtgdvwClientes.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                      Convert.ToInt32(dtgdvwClientes.CurrentRow.Cells["Ativos"].EditedFormattedValue.ToString()));
                }
                bolAddBd = false;
            }
            else
            {
                if (MessageBox.Show("Deseja Alterar" + strValorAntigo + "Para" + dtgdvwClientes.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString() + "? ", "Alterando Registro", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    AlterarBD(intValorSalvo, dtgdvwClientes.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(),
                     dtgdvwClientes.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                     Convert.ToInt32(dtgdvwClientes.CurrentRow.Cells["Ativos"].EditedFormattedValue.ToString()));
                }
            }
            ConsultarBD();
        }

        private void bndnavbtnPesquisarCL_Click(object sender, EventArgs e)
        {
            ConsultarBD(null, bndnavtxtPesquisarCL.Text);
        }

        private void dtgdvwClientes_CellClick(object sender, DataGridViewCellEventArgs e)
        {
                if (!string.IsNullOrEmpty(dtgdvwClientes.CurrentRow.Cells["ID"].Value.ToString()))
                {
                    intValorSalvo = Convert.ToInt32(dtgdvwClientes.CurrentRow.Cells["ID"].Value.ToString());
                }
                strValorAntigo = dtgdvwClientes.CurrentRow.Cells["Nome"].Value.ToString();
        }
        #endregion
        #region Pedidos Interior
        public void ConsultarBDPI(int? intID = null,
                               int? intCliente_ID = null,
                               string strNome = null,
                               string strDescricao = null,
                               int? intEstado = null)
        {
            try
            {
                objPedidos_Interior_BLL = new Pedidos_Interior_BLL();
                objPedidos_Interior_VO = new Pedidos_Interior_VO();
                objPedidos_Interior_VO.ID = Convert.ToInt32(intID == null ? 0 : intID);

                objPedidos_Interior_VO.Cliente_ID = new Cliente_VO();
                objPedidos_Interior_VO.Cliente_ID.ID = Convert.ToInt32(intCliente_ID == null ? 0 : intCliente_ID);
                objPedidos_Interior_VO.Cliente_ID.Nome = strNome;

                bndsrcPedidosInterior.DataSource = objPedidos_Interior_BLL.ConsultarBD(objPedidos_Interior_VO);

                dtgdvwPedidosInterior.DataSource = null;
                dtgdvwPedidosInterior.Columns.Clear();
                dtgdvwPedidosInterior.AllowUserToAddRows = false;

                dtgdvwPedidosInterior.Columns.Add("ID", "ID do Pedido");
                dtgdvwPedidosInterior.Columns["ID"].DataPropertyName = "ID";

                DataGridViewComboBoxColumn objdtgdvwcmbbxPI = new DataGridViewComboBoxColumn();
                objdtgdvwcmbbxPI.DataSource = bndsrcClientes.DataSource;
                objdtgdvwcmbbxPI.Name = "Cliente_ID";
                objdtgdvwcmbbxPI.ValueType = typeof(int);
                objdtgdvwcmbbxPI.DisplayMember = "Nome";
                objdtgdvwcmbbxPI.ValueMember = "ID";
                objdtgdvwcmbbxPI.HeaderText = "Nome de Cliente";

                dtgdvwPedidosInterior.Columns.Add(objdtgdvwcmbbxPI);
                dtgdvwPedidosInterior.Columns["Cliente_ID"].DataPropertyName = "Cliente_ID";

                dtgdvwPedidosInterior.Columns.Add("Descricao", "Descricao do Pedido");
                dtgdvwPedidosInterior.Columns["Descricao"].DataPropertyName = "Descricao";

                dtgdvwPedidosInterior.Columns.Add("Estado", "Estado do Pedido");
                dtgdvwPedidosInterior.Columns["Estado"].DataPropertyName = "Estado";

                dtgdvwPedidosInterior.DataSource = bndsrcPedidosInterior;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Presta Atenção !!!  :" + ex.Message);
            }
        }
        public void IncluirBDPI(int intCliente_ID, string strNome, string strDescricao, int intAtivo)
        {
            try
            {
                objPedidos_Interior_BLL = new Pedidos_Interior_BLL();
                objPedidos_Interior_VO = new Pedidos_Interior_VO();

                objPedidos_Interior_VO.Cliente_ID = new Cliente_VO();

                objPedidos_Interior_VO.Cliente_ID.ID = intCliente_ID <= 0 ? 0 : intCliente_ID;
                objPedidos_Interior_VO.Cliente_ID.Nome = strNome;
                objPedidos_Interior_VO.Descricao = strDescricao;
                objPedidos_Interior_VO.Cliente_ID.Ativos = intAtivo;

                if (objPedidos_Interior_BLL.IncluirBD(objPedidos_Interior_VO))
                {
                    MessageBox.Show("Registro Incluso!");
                }
                else
                {
                    MessageBox.Show("Registro Não Incluso!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Presta Atenção !!!  :" + ex.Message);
            }
        }
        public void ExcluirBDPI(int intID)
        {
            try
            {
                objPedidos_Interior_BLL = new Pedidos_Interior_BLL();
                objPedidos_Interior_VO = new Pedidos_Interior_VO();

                objPedidos_Interior_VO.ID = intID;


                if (objPedidos_Interior_BLL.ExcluirBD(objPedidos_Interior_VO))
                {
                    MessageBox.Show("Registro Excluido!");
                }
                else
                {
                    MessageBox.Show("Registro Não Excluido!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Presta Atenção !!!  :" + ex.Message);
            }
        }

        public void AlterarBDPI(int intID, int intCliente_ID, string strNome, string strDescricao, int intAtivo)
        {
            try
            {
                objPedidos_Interior_BLL = new Pedidos_Interior_BLL();
                objPedidos_Interior_VO = new Pedidos_Interior_VO();

                objPedidos_Interior_VO.ID = intID;
                objPedidos_Interior_VO.Cliente_ID.Nome = strNome;
                objPedidos_Interior_VO.Descricao = strDescricao;
                objPedidos_Interior_VO.Cliente_ID.Ativos = intAtivo;

                if (objPedidos_Interior_BLL.AlterarBD(objPedidos_Interior_VO))
                {
                    MessageBox.Show("Registro Alterado!");
                }
                else
                {
                    MessageBox.Show("Registro Não Alterado!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Presta Atenção !!!  :" + ex.Message);
            }
        }

        public void dtgdvw_PI_Refresh()
        {
            ConsultarBDPI(null,Convert.ToInt32(cmbbxPI.SelectedValue.ToString()),cmbbxPI.Text);
        }
        private void cmbbxPI_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbbxPI.SelectedIndex >= 0)
            {
                cmbbxPI.Text = cmbbxPI.Text.Trim();
            }
        }
        private void cmbbxPI_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(cmbbxPI.Text))
            {
                strNome_Salvo_PI = cmbbxPI.Text;
                dtgdvw_PI_Refresh();
            }
        }
        public void Reconfigurar_Grid_Interior(bool bolSelectedClientes_ID, bool bolReadOnlyClientes_ID)
        {
            dtgdvwPedidosInterior.CurrentRow.Cells["Cliente_ID"].Selected = bolSelectedClientes_ID;
            dtgdvwPedidosInterior.CurrentRow.Cells["Cliente_ID"].ReadOnly = bolReadOnlyClientes_ID;
        }
        private void dtgdvwPI_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!string.IsNullOrEmpty(dtgdvwPedidosInterior.CurrentRow.Cells["ID"].Value.ToString()) && 
                !string.IsNullOrEmpty(dtgdvwPedidosInterior.CurrentRow.Cells["Cliente_ID"].Value.ToString()))
            {
                intID_Salvo_PI = Convert.ToInt32(dtgdvwPedidosInterior.CurrentRow.Cells["ID"].Value.ToString());
                intCliente_ID_Salvo_PI = Convert.ToInt32(dtgdvwPedidosInterior.CurrentRow.Cells["Cliente_ID"].Value.ToString());
                strNome_Salvo_PI = dtgdvwPedidosInterior.CurrentRow.Cells["Cliente_ID"].EditedFormattedValue.ToString();
                Reconfigurar_Grid_Interior(false,true);
            }
            else
            {
                dtgdvwPedidosInterior.CurrentRow.Cells["Cliente_ID"].Value = cmbbxPI.SelectedValue;
                Reconfigurar_Grid_Interior(false, true);
            }
        }
        private void bndnavbtnPIAdd_Click(object sender, EventArgs e)
        {
            bolAddBd_PI = true;
            Reconfigurar_Grid_Interior(false, true);
        }
        private void bndnavbtnPIAExcluir_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Excluir" + strNome_Salvo_PI + "? ", "Excluir Registro  ", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                ExcluirBDPI(intID_Salvo_PI);
            }
            dtgdvw_PI_Refresh();
        }
        private void bndnavbtnPIAConfirmar_Click(object sender, EventArgs e)
        {
            if (bolAddBd_PI)
            {
                if (MessageBox.Show("Deseja Incluir" + dtgdvwPedidosInterior.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString() + "? ", "Incluindo Registro  ", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    IncluirBDPI(Convert.ToInt32(dtgdvwPedidosInterior.CurrentRow.Cells["Cliente_ID"].Value),
                                                dtgdvwPedidosInterior.CurrentRow.Cells["Cliente_ID"].Value.ToString(),
                                                dtgdvwPedidosInterior.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                                Convert.ToInt32(dtgdvwPedidosInterior.CurrentRow.Cells["Estado"].EditedFormattedValue.ToString()));
                }
                bolAddBd_PI = false;
                Reconfigurar_Grid_Interior(false,true);
            }
            else
            {
                if (MessageBox.Show("Deseja Alterar" + strNome_Salvo_PI + "Para " + dtgdvwPedidosInterior.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString() + "? ", "Incluindo Registro  ", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    AlterarBDPI(intID_Salvo_PI,
                                intCliente_ID_Salvo_PI,
                                strNome_Salvo_PI,
                                dtgdvwPedidosInterior.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                                Convert.ToInt32(dtgdvwPedidosInterior.CurrentRow.Cells["Estado"].EditedFormattedValue.ToString()));
                }
            }
            dtgdvw_PI_Refresh();
        }
        private void bndnavbtnPIPesquisar_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(bndnavtxtPIPesqui.Text))
                {
                    ConsultarBDPI(Convert.ToInt32(bndnavtxtPIPesqui.Text));
                }
                else
                {
                    dtgdvw_PI_Refresh();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        #region Pedidos Exterior
        public void ConsultarBDPE(int? intID = null, int? intCliente_ID = null, string strNome = null, string strDescricao = null, int? intEstado = null)
        {
            try
            {
                objPedidos_Exterior_BLL = new Pedidos_Exterior_BLL();
                objPedidos_Exterior_VO = new Pedidos_Exterior_VO();
                objPedidos_Exterior_VO.ID = Convert.ToInt32(intID == null ? 0 : intID);

                objPedidos_Exterior_VO.Cliente_ID = new Cliente_VO();
                objPedidos_Exterior_VO.Cliente_ID.ID = Convert.ToInt32(intCliente_ID == null ? 0 : intCliente_ID);
                objPedidos_Exterior_VO.Cliente_ID.Nome = strNome;

                bndsrcPedidosExterior.DataSource = objPedidos_Exterior_BLL.ConsultarBD(objPedidos_Exterior_VO);

                dtgdvwPedidosExterior.DataSource = null;
                dtgdvwPedidosExterior.Columns.Clear();
                dtgdvwPedidosExterior.AllowUserToAddRows = false;

                dtgdvwPedidosExterior.Columns.Add("ID", "ID do Pedido");
                dtgdvwPedidosExterior.Columns["ID"].DataPropertyName = "ID";

                DataGridViewComboBoxColumn objdtgdvwcmbbxPE = new DataGridViewComboBoxColumn();
                objdtgdvwcmbbxPE.DataSource = bndsrcClientes.DataSource;
                objdtgdvwcmbbxPE.Name = "Cliente_ID";
                objdtgdvwcmbbxPE.ValueType = typeof(int);
                objdtgdvwcmbbxPE.DisplayMember = "Nome";
                objdtgdvwcmbbxPE.ValueMember = "ID";
                objdtgdvwcmbbxPE.HeaderText = "Nome de Cliente";

                dtgdvwPedidosExterior.Columns.Add(objdtgdvwcmbbxPE);
                dtgdvwPedidosExterior.Columns["Cliente_ID"].DataPropertyName = "Cliente_ID";

                dtgdvwPedidosExterior.Columns.Add("Descricao", "Descricao do Pedido");
                dtgdvwPedidosExterior.Columns["Descricao"].DataPropertyName = "Descricao";

                dtgdvwPedidosExterior.Columns.Add("Estado", "Estado do Pedido");
                dtgdvwPedidosExterior.Columns["Estado"].DataPropertyName = "Estado";

                dtgdvwPedidosExterior.DataSource = bndsrcPedidosExterior;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Presta Atenção !!!  :" + ex.Message);
            }
        }
        public void IncluirBDPE(int intCliente_ID, string strNome, string strDescricao, int intAtivo)
        {
            try
            {
                objPedidos_Exterior_BLL = new Pedidos_Exterior_BLL();
                objPedidos_Exterior_VO = new Pedidos_Exterior_VO();

                objPedidos_Exterior_VO.Cliente_ID = new Cliente_VO();

                objPedidos_Exterior_VO.Cliente_ID.ID = intCliente_ID <= 0 ? 0 : intCliente_ID;
                objPedidos_Exterior_VO.Cliente_ID.Nome = strNome;
                objPedidos_Exterior_VO.Descricao = strDescricao;
                objPedidos_Exterior_VO.Cliente_ID.Ativos = intAtivo;

                if (objPedidos_Exterior_BLL.IncluirBD(objPedidos_Exterior_VO))
                {
                    MessageBox.Show("Registro Incluso!");
                }
                else
                {
                    MessageBox.Show("Registro Não Incluso!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Presta Atenção !!!  :" + ex.Message);
            }
        }
        public void ExcluirBDPE(int intID)
        {
            try
            {
                objPedidos_Exterior_BLL = new Pedidos_Exterior_BLL();
                objPedidos_Exterior_VO = new Pedidos_Exterior_VO();

                objPedidos_Exterior_VO.ID = intID;


                if (objPedidos_Exterior_BLL.ExcluirBD(objPedidos_Exterior_VO))
                {
                    MessageBox.Show("Registro Excluido!");
                }
                else
                {
                    MessageBox.Show("Registro Não Excluido!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Presta Atenção !!!  :" + ex.Message);
            }
        }
        public void AlterarBDPE(int intID, int intCliente_ID, string strNome, string strDescricao, int intAtivo)
        {
            try
            {
                objPedidos_Exterior_BLL = new Pedidos_Exterior_BLL();
                objPedidos_Exterior_VO = new Pedidos_Exterior_VO();

                objPedidos_Exterior_VO.ID = intID;
                objPedidos_Exterior_VO.Cliente_ID.Nome = strNome;
                objPedidos_Exterior_VO.Descricao = strDescricao;
                objPedidos_Exterior_VO.Cliente_ID.Ativos = intAtivo;

                if (objPedidos_Exterior_BLL.AlterarBD(objPedidos_Exterior_VO))
                {
                    MessageBox.Show("Registro Alterado!");
                }
                else
                {
                    MessageBox.Show("Registro Não Alterado!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Presta Atenção !!!  :" + ex.Message);
            }
        }
        public void dtgdvw_PE_Refresh()
        {
            ConsultarBDPE(null, Convert.ToInt32(cmbbxPE.SelectedValue.ToString()),cmbbxPE.Text);
        }
        private void cmbbxPE_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbbxPE.SelectedIndex >= 0)
            {
                cmbbxPE.Text = cmbbxPE.Text.Trim();
            }
        }
        private void cmbbxPE_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(cmbbxPE.Text))
            {
                strNome_Salvo_PE = cmbbxPE.Text;
                dtgdvw_PE_Refresh();
            }
        }
        public void Reconfigurar_Grid_Exterior(bool bolSelectedClientes_ID, bool bolReadOnlyClientes_ID)
        {
            dtgdvwPedidosExterior.CurrentRow.Cells["Cliente_ID"].Selected = bolSelectedClientes_ID;
            dtgdvwPedidosExterior.CurrentRow.Cells["Cliente_ID"].ReadOnly = bolReadOnlyClientes_ID;
        }
        private void dtgdvwPedidosExterior_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!string.IsNullOrEmpty(dtgdvwPedidosExterior.CurrentRow.Cells["ID"].Value.ToString()) &&
                !string.IsNullOrEmpty(dtgdvwPedidosExterior.CurrentRow.Cells["Cliente_ID"].Value.ToString()))
            {
                intID_Salvo_PE = Convert.ToInt32(dtgdvwPedidosExterior.CurrentRow.Cells["ID"].Value.ToString());
                intCliente_ID_Salvo_PE = Convert.ToInt32(dtgdvwPedidosExterior.CurrentRow.Cells["Cliente_ID"].Value.ToString());
                strNome_Salvo_PE = dtgdvwPedidosExterior.CurrentRow.Cells["Cliente_ID"].EditedFormattedValue.ToString();
                Reconfigurar_Grid_Exterior(false, true);
            }
            else
            {
                dtgdvwPedidosExterior.CurrentRow.Cells["Cliente_ID"].Value = cmbbxPE.SelectedValue;
                Reconfigurar_Grid_Exterior(false, true);
            }
        }
        private void bndnavbtnAdd_Click(object sender, EventArgs e)
        {
            bolAddBd_PE = true;
            Reconfigurar_Grid_Exterior(false, true);
        }

        private void bndnavbtnExcl_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Excluir" + strNome_Salvo_PE + "? ", "Excluir Registro  ", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                ExcluirBDPE(intID_Salvo_PE);
            }
            dtgdvw_PE_Refresh();
        }

        private void bndnavbtnPeConirmar_Click(object sender, EventArgs e)
        {
            if (bolAddBd_PE)
            {
                if (MessageBox.Show("Deseja Incluir" + dtgdvwPedidosExterior.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString() + "? ", "Incluindo Registro  ", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    IncluirBDPE(Convert.ToInt32(dtgdvwPedidosExterior.CurrentRow.Cells["Cliente_ID"].Value),
                                                dtgdvwPedidosExterior.CurrentRow.Cells["Cliente_ID"].Value.ToString(),
                                                dtgdvwPedidosExterior.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                                Convert.ToInt32(dtgdvwPedidosExterior.CurrentRow.Cells["Estado"].EditedFormattedValue.ToString()));
                }
                bolAddBd_PE = false;
                Reconfigurar_Grid_Exterior(false, true);
            }
            else
            {
                if (MessageBox.Show("Deseja Alterar" + strNome_Salvo_PE + "Para " + dtgdvwPedidosExterior.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString() + "? ", "Incluindo Registro  ", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    AlterarBDPE(intID_Salvo_PE,
                                intCliente_ID_Salvo_PE,
                                strNome_Salvo_PE,
                                dtgdvwPedidosExterior.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                                Convert.ToInt32(dtgdvwPedidosExterior.CurrentRow.Cells["Estado"].EditedFormattedValue.ToString()));
                }
            }
            dtgdvw_PE_Refresh();
        }

        private void bndnavbtnPePesquiar_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(bndnavtxtPesquisar.Text))
                {
                    ConsultarBDPE(Convert.ToInt32(bndnavtxtPesquisar.Text));
                }
                else
                {
                    dtgdvw_PE_Refresh();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        #region Consultas
        private void btnConsultar_Clientes_Sem_Pedidos_Exterior_Click(object sender, EventArgs e)
        {
            Consulta1();
        }

        public void Consulta1()
        {
            objCliente_BLL = new Cliente_BLL();
            bndsrcConsulta1.DataSource = objCliente_BLL.Consultar_Clientes_Sem_Pedidos_Exterior();
            dtgdvwConsulta1.DataSource = bndsrcConsulta1;
        }

        private void btnConsultar_Pedidos_De_Clientes_Slecionados_Click(object sender, EventArgs e)
        {
            strPesquisarIdClientes = Interaction.InputBox("Digite o ID dos Clientes", "Consulta de ID");
            Consulta2();
        }
        public void Consulta2()
        {
            objCliente_BLL = new Cliente_BLL();
            bndsrcConsulta2.DataSource = objCliente_BLL.Consultar_Pedidos_De_Clientes_Slecionados(strPesquisarIdClientes);
            dtgdvwConsulta2.DataSource = bndsrcConsulta2;
        }
        private void btnConsultar_De_Quantidade_De_PedidosInterior_Por_Clientes_Click(object sender, EventArgs e)
        {
            strQuantidade_de_PI_Cliente = Interaction.InputBox("Digite o ID dos Clientes", "Consulta de ID");
            Consulta3();
        }
        public void Consulta3()
        {
            objCliente_BLL = new Cliente_BLL();
            bndsrcConsulta3.DataSource = objCliente_BLL.Consultar_De_Quantidade_De_PedidosInterior_Por_Clientes(strQuantidade_de_PI_Cliente);
            dtgdvwConsulta3.DataSource = bndsrcConsulta3;
        }

        private void btnConsultar_Quantidades_De_Pedidos_Dos_Clientes_Click(object sender, EventArgs e)
        {
            Consulta4();
        }
        public void Consulta4()
        {
            objCliente_BLL = new Cliente_BLL();
            bndsrcConsulta4.DataSource = objCliente_BLL.Consultar_Quantidades_De_Pedidos_Dos_Clientes();
            dtgdvwConsulta4.DataSource = bndsrcConsulta4;
        }
        #endregion
        #region Automacao Excel
        public void Automacao_Excel_BD(DataTable objTabelaExcel)
        {
            try
            {
                if (objTabelaExcel != null)
                {
                    objApplication = new Excel.Application();
                    objApplication.Visible = true;
                    objWorkbook = objApplication.Workbooks.Add();
                    objWorksheet = objWorkbook.Worksheets[1];

                    int intColuna = 1, intLinha = 2, intLinhaCabecalho = 1;

                    objCabecalho = objWorksheet.Cells[intLinhaCabecalho, intColuna];
                    objExDados = objWorksheet.Cells[intLinha, intColuna];

                    foreach (DataRow objLinhaBD in objTabelaExcel.Rows)
                    {
                        foreach (DataColumn objColunaBD in objTabelaExcel.Columns)
                        {
                            if (intLinha <= intLinhaCabecalho + 1)
                            {
                                objCabecalho.set_Value(Type.Missing, objColunaBD.ColumnName);
                            }

                            if (!string.IsNullOrEmpty(objLinhaBD[intColuna - 1].ToString()))
                            {
                                objExDados.set_Value(Type.Missing, objLinhaBD[intColuna - 1].ToString());
                            }

                            intColuna++;

                            if (intLinha <= intLinhaCabecalho + 1)
                            {
                                objCabecalho = objWorksheet.Cells[intLinhaCabecalho, intColuna];
                            }
                            objExDados = objWorksheet.Cells[intLinha, intColuna];
                        }

                        intLinha++;
                        intColuna = 1;
                        objExDados = objWorksheet.Cells[intLinha, intColuna];
                    }
                    sfdSalveExcel.ShowDialog();
                    objWorksheet.SaveAs(sfdSalveExcel.FileName.ToString(), Type.Missing,
                                                                          Type.Missing,
                                                                          Type.Missing,
                                                                          Type.Missing,
                                                                          Type.Missing, Excel.XlSaveAsAccessMode.xlShared);
                }
                else
                {
                    MessageBox.Show("Faça a consulta antes de abrir com Excel", "Erro Excel", MessageBoxButtons.OK,MessageBoxIcon.Warning);
                }
               

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }
        private void bndnavbtnGerarExelCliente_Click(object sender, EventArgs e)
        {
            Automacao_Excel_BD((DataTable)bndsrcClientes.DataSource);
        }

        private void bndnavbtnGerarExelPI_Click(object sender, EventArgs e)
        {
            Automacao_Excel_BD((DataTable)bndsrcPedidosInterior.DataSource);
        }

        private void bndnavbtnGerarExelPE_Click(object sender, EventArgs e)
        {
            Automacao_Excel_BD((DataTable)bndsrcPedidosExterior.DataSource);
        }

        private void bndnavbtnGerarExelConsulta1_Click(object sender, EventArgs e)
        {
            Automacao_Excel_BD((DataTable)bndsrcConsulta1.DataSource);
        }

        private void bndnavbtnGerarExelConsulta2_Click(object sender, EventArgs e)
        {
            Automacao_Excel_BD((DataTable)bndsrcConsulta2.DataSource);
        }

        private void bndnavbtnGerarExelConsulta3_Click(object sender, EventArgs e)
        {
            Automacao_Excel_BD((DataTable)bndsrcConsulta3.DataSource);
        }

        private void bndnavbtnGerarExelConsulta4_Click(object sender, EventArgs e)
        {
            Automacao_Excel_BD((DataTable)bndsrcConsulta4.DataSource);
        }
        #endregion
        #region Automacao Email
        private void bndnavbtnGerarEmail_Click(object sender, EventArgs e)
        {
            GerarEmailOutlook();
        }
        public void GerarEmailOutlook()
        {
            objEmailApp = new Email.Application();
            objEmailMsn = objEmailApp.CreateItem(Email.OlItemType.olMailItem);
            objEmailMsn.SentOnBehalfOfName = "adriano.ssud.Cordeiro@uotlook.com";
            objEmailMsn.To = "adr.sud.cor@gmail.com";
            objEmailMsn.Subject = "Sistema de Cadastros de Pedidos e Clientes ";
            objEmailMsn.Body = "Bom dia,\nSegue em anexo a tabela de Clientes no formato .xlsx";

            if (MessageBox.Show("Deseja anexar os arquivos?", "Arquivos anexo", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                ofdArqAnexo.Title = "Escolha os Arquivos Para Anexo";
                ofdArqAnexo.InitialDirectory = @"C:\Atomacao_Bancaria\Arquivos Excel";
                ofdArqAnexo.ShowDialog();

                string strEnderecoDeAnexo = ofdArqAnexo.FileName;

                if (!string.IsNullOrEmpty(strEnderecoDeAnexo))
                {
                    for (int z = 0; z < objAnexoArq.Length; z++)
                    {
                        objAttchment = Email.OlAttachmentType.olByValue;
                        objAnexoPosition = objEmailMsn.Body.Length + 1;
                        objDisplayName = objAnexoArq[z].ToString() + "- novo arquivo - Email Excel";
                        objEmailMsn.Attachments.Add(objAnexoArq[z], objAttchment, objAnexoPosition, objDisplayName);
                    }
                }
            }
            if (MessageBox.Show("Enviar o Email com confirmação ? ","Confirmação de Envio", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                objEmailMsn.Display();
            }
            else
            {
                objEmailMsn.Send();
            }
            MessageBox.Show("Email Enviado com Sucesso !!!", "Envio de Email", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion

        private void frmAltomacaoFull_Load(object sender, EventArgs e)
        {
            ConsultarBD();
            dtgdvw_PI_Refresh();
            dtgdvw_PE_Refresh();
        }

       
    }
}
