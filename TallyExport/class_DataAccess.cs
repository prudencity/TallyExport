using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace TallyExport
{
    public class class_DataAccess : IDisposable
    {
        private bool disposedValue = false;
        public SqlConnection Conn;

        public void CreateConnection()
        {
            if (Conn == null)
            {
                Conn = new SqlConnection(Program.connectionString);                     //Connection string stored here
            }
            if (Conn.State != ConnectionState.Open)
            {
                Conn.Open();
            }
        }

        public SqlConnection GetConnectionObj()
        {
            CreateConnection();
            return Conn;
        }

        public void CloseConnection()
        {
            if (Conn != null)
            {
                if (Conn.State != ConnectionState.Closed)
                {
                    Conn.Close();
                }

                Conn = null;
            }
        }

        public DataSet GetDataSet(string strsql)
        {
            CreateConnection();
            SqlDataAdapter da = new SqlDataAdapter(strsql, Conn);
            DataSet ds = new DataSet();
            da.Fill(ds);
            return ds;
        }

        public DataTable GetDataTable(string strsql)
        {
            CreateConnection();
            SqlDataAdapter da = new SqlDataAdapter(strsql, Conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }

        public SqlDataReader GetReader(string strsql)
        {
            CreateConnection();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strsql;
            cmd.Connection = Conn;
            SqlDataReader reader = cmd.ExecuteReader();
            return reader;
        }

        public bool reader(string strsql)
        {
            CreateConnection();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strsql;
            SqlDataReader reader = cmd.ExecuteReader();
            bool ans = reader.Read();
            reader.Close();
            return ans;
        }

        public int ExecuteQuery(string strsql)
        {
            CreateConnection();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = Conn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strsql;
            int i;
            i = cmd.ExecuteNonQuery();
            return i;
        }

        public int ExecuteQueryIdentity(string strsql)
        {
            CreateConnection();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = Conn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strsql + "; select scope_identity();";
            int i;
            i = Convert.ToInt32(cmd.ExecuteScalar());
            return i;
        }

        public int GetInsertIndentity(string strsql)
        {
            CreateConnection();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = Conn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strsql + "; select scope_identity();";
            int i;
            i = Convert.ToInt32(cmd.ExecuteScalar());
            return i;
        }

        #region " IDisposable Support "
        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposedValue)
            {
                if (disposing)
                {
                    CloseConnection();

                }
            }
            this.disposedValue = true;
        }

        public void Dispose()
        {
            if (!(Conn.State == ConnectionState.Closed))
            {
                CloseConnection();
            }
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion
    }
}
