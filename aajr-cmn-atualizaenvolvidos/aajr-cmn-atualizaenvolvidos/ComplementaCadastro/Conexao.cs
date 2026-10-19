using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ComplementaCadastro
{
    static class Conexao
    {
        public static string con = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\\BD\\BD_DADOS.accdb";
        public static OleDbConnection conexao;
        public static OleDbDataReader da;
        public static OleDbCommand cmd;


        public static OleDbConnection conectar()
        {
            conexao = new OleDbConnection(con);
            conexao.Open();
            return conexao;
        }

        public static String bancoAtual()
        {
            return System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\\BD_DADOS.accdb";
        }

        public static void desconectar()
        {
            conexao.Close();
        }

        public static OleDbDataReader consultar(string comando)
        {
            cmd = new OleDbCommand();
            cmd.Connection = conectar();
            cmd.CommandText = comando;
            try 
	        {	        
		        da = cmd.ExecuteReader();
	        }
	        catch (Exception ex)
	        {
		
		        throw ex;
	        }
            return da;

        }

        public static DataTable GeraTBL(string comando)
        {
            cmd = new OleDbCommand();
            cmd.Connection = conectar();
            cmd.CommandText = comando;

            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataTable dt = new DataTable();

            try
            {
                da.Fill(dt);
            }
            catch (Exception ex)
            {

                throw ex;
            }
            // desconectar();
            return dt;

        }

        public static Boolean executar(string comando)
        {
            cmd = new OleDbCommand();
            cmd.Connection = conectar();
            cmd.CommandText = comando;
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());
            }
            finally
            {
                desconectar();
            }
            return true;
        }

        public static Boolean executarProcedure(string comando)
        {
            cmd = new OleDbCommand();
            cmd.CommandType = System.Data.CommandType.StoredProcedure;
            cmd.Connection = conectar();
            cmd.CommandText = comando;
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception)
            {
                desconectar();
                return false;
            }

            desconectar();
            return true;
        }

    }
}
