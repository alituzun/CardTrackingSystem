using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApplication17
{
    public class DAL  
    {  
        public static int executeQuery(string query)  
        {  
            int rowCount = 0;
            string strConn = @"Data Source=(LocalDB)\v11.0;AttachDbFilename=D:\Users\st900398\Documents\deneme.mdf;Integrated Security=True;Connect Timeout=30";  
            SqlConnection sqlConnection = new SqlConnection(strConn);  
           SqlCommand sqlCommand = new SqlCommand();  
            try  
            {   
                sqlCommand.CommandText = query;  
                sqlCommand.CommandType = CommandType.Text;  
                sqlCommand.Connection = sqlConnection;  
                sqlConnection.Open();  
                rowCount = sqlCommand.ExecuteNonQuery();  
                sqlConnection.Close();  
            }  
            catch (Exception ex)  
            {  
                sqlConnection.Close();  
                throw ex;  
            }  
            return rowCount;  
        }  
    }  
}
