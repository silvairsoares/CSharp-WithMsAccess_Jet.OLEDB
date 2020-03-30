using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace AcessoBD_Dicas
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ExibirDados();
        }

        private void ExibirDados()
        {
            // definir a string de conexão
            string sDBstr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dados\Teste.mdb";

            //definir a string SQL
            string sSQL = "SELECT * from Clientes";

            //criar o objeto connection
            OleDbConnection oCn = new OleDbConnection(sDBstr);
            //abrir a conexão
            oCn.Open();
            //criar o data adapter e executar a consulta 
            OleDbDataAdapter oDA = new OleDbDataAdapter(sSQL, oCn);
            //criar o DataSet
            DataSet oDs = new DataSet();
            //Preencher o dataset coom o data adapter
            oDA.Fill(oDs, "Clientes");

            //exibe os dados no datagridview
            //exibe o resultado
            dgvDados.DataSource = oDs.Tables[0];

            // liberar o data adapter , o dataset e a conexao
            oDA.Dispose();
            oDs.Dispose();
            oCn.Dispose();
        }

        private void btnInserir_Click(object sender, EventArgs e)
        {
            // definir a string de conexão
            string sDBstr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dados\Teste.mdb";

            //definir a string SQL
            string sSQL = "SELECT * from Clientes";

            //criar o objeto connection
            OleDbConnection oCn = new OleDbConnection(sDBstr);
            //abrir a conexão
            oCn.Open();
            //criar o data adapter e executar a consulta 
            OleDbDataAdapter oDA = new OleDbDataAdapter(sSQL, oCn);
            //criar o DataSet
            DataSet oDs = new DataSet();
            //Preencher o dataset coom o data adapter
            oDA.Fill(oDs, "Clientes");

            //criar um objeto Data Row
            DataRow oDR = oDs.Tables["Clientes"].NewRow();

            //Preencher o datarow com valores
            oDR["Nome"] = "Testolino";
            oDR["Email"] = "teste@teste.com.br";
         
            //Incluir um datarow ao dataset
            oDs.Tables["Clientes"].Rows.Add(oDR);
            //Usar o objeto Command Bulder para gerar o Comandop Insert 
            OleDbCommandBuilder oCB = new OleDbCommandBuilder(oDA);
            //Atualizar o BD com valores do Dataset 
            oDA.Update(oDs, "Clientes");

            //liberar o data adapter , o dataset , o comandbuilder e a conexao
            oDA.Dispose();
            oDs.Dispose();
            oCB.Dispose();
            oCn.Dispose();

            //exibir o resultado
            ExibirDados();
        }

       

        private void btnAtualizar_Click(object sender, EventArgs e)
        {
            // definir a string de conexão
            string sDBstr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dados\Teste.mdb";

            //definir a string SQL
            string sSQL = "SELECT * from Clientes";

            //criar o objeto connection
            OleDbConnection oCn = new OleDbConnection(sDBstr);
            //abrir a conexão
            oCn.Open();
            //criar o data adapter e executar a consulta 
            OleDbDataAdapter oDA = new OleDbDataAdapter(sSQL, oCn);
            //criar o DataSet
            DataSet oDs = new DataSet();
            //Preencher o dataset coom o data adapter
            oDA.Fill(oDs, "Clientes");
            //cria o DataSet atribuindo ao DataRow o valor da linha que desejamos atualizar
            DataRow oDR = oDs.Tables["Clientes"].Rows[5];

            //Preenche o datarow with valores
            oDR["Nome"] = "Maria";
            oDR["Email"] = "maria@net.com.br";
            
            //Usar o objeto Command Bulder para gerar o Comando Update
            OleDbCommandBuilder oCB = new OleDbCommandBuilder(oDA);
            //Atualizar o BD com valores do Dataset 
            oDA.Update(oDs, "Clientes");

            //liberar o data adapter , o dataset , o comandbuilder e a conexao
            oDA.Dispose();
            oDs.Dispose();
            oCB.Dispose();
            oCn.Dispose();

            //exibir o resultado
            ExibirDados();
        }

        private void btnExcluir_Click(object sender, EventArgs e)
        {
            // definir a string de conexão
            string sDBstr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dados\Teste.mdb";

            //definir a string SQL
            string sSQL = "SELECT * from Clientes";

            //criar o objeto connection
            OleDbConnection oCn = new OleDbConnection(sDBstr);
            //abrir a conexão
            oCn.Open();
            //criar o data adapter e executar a consulta 
            OleDbDataAdapter oDA = new OleDbDataAdapter(sSQL, oCn);
            //criar o DataSet
            DataSet oDs = new DataSet();
            //Preencher o dataset coom o data adapter
            oDA.Fill(oDs, "Clientes");
            //Exclui a linha desejada
            oDs.Tables["Clientes"].Rows[5].Delete();

            // Usar o objeto Command Bulder para gerar o Comandop Delete
            OleDbCommandBuilder oCB = new OleDbCommandBuilder(oDA);
            //Atualizar o BD com valores do Dataset 
            oDA.Update(oDs, "Clientes");

            //liberar o data adapter , o dataset , o comandbuilder e a conexao
            oDA.Dispose();
            oDs.Dispose();
            oCB.Dispose();
            oCn.Dispose();

            //exibir o resultado
            ExibirDados();
        }
    }
}
