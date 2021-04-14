using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ExampleCS
{
    struct Field
    {
        public int id;
        public string name;
    }

    class Utils
    {
        private List<Field> fields;

        public Utils()
        {
            fields = new List<Field>();
        }

        public void Download()
        {
            OleDbConnection dbConnection = new OleDbConnection();

            try
            {
                // Creating a connection
                string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";
                dbConnection = new OleDbConnection(connectionString);

                // Executing a query to the database
                dbConnection.Open();
                string query = "SELECT * FROM table_name";
                OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
                OleDbDataReader dbReader = dbCommand.ExecuteReader(); // reading the data

                // Checking the data
                if (dbReader.HasRows == false)
                {
                    MessageBox.Show("Данные не найдены!", "Ошибка!");
                }
                else
                {
                    fields.Clear();

                    // Reading the data
                    while (dbReader.Read())
                    {
                        Field buf = new Field();
                        buf.id = Convert.ToInt32(dbReader["id"]);
                        buf.name = Convert.ToString(dbReader["name"]);

                        fields.Add(buf);
                    }
                }

                dbReader.Close();
            }
            catch (OleDbException e)
            {
                MessageBox.Show(e.ToString(),"Error");
            }
            finally
            {
                // Closing the connection to the database
                dbConnection.Close();
            }
        }

        public void Insert()
        {
            OleDbConnection dbConnection = new OleDbConnection();

            try
            {
                // Creating a connection
                string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";
                dbConnection = new OleDbConnection(connectionString);

                // Executing a query to the database
                dbConnection.Open();

                // Insert all data
                bool err = false;
                for(int i = 0; i<fields.Count; i++)
                {
                    string query = "INSERT INTO table_name VALUES (" + fields[i].id + ", '" + fields[i].name + ")";
                    OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);

                    // Executing the request
                    if (dbCommand.ExecuteNonQuery() != 1)
                        err = true;
                }

                if(err)
                    MessageBox.Show("Inserting error!", "Error!");
                else
                    MessageBox.Show("Data inserted!", "Done!");

            }
            catch (OleDbException e)
            {
                MessageBox.Show(e.ToString(), "Error");
            }
            finally
            {
                // Closing the connection to the database
                dbConnection.Close();
            }
        }

        public void Insert(Field f)
        {
            OleDbConnection dbConnection = new OleDbConnection();

            try
            {
                // Creating a connection
                string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";
                dbConnection = new OleDbConnection(connectionString);

                // Executing a query to the database
                dbConnection.Open();
                string query = "INSERT INTO table_name VALUES (" + f.id + ", '" + f.name + ")";
                OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);

                // Executing the request
                if (dbCommand.ExecuteNonQuery() != 1)
                    MessageBox.Show("Inserting error!", "Error!");
                else
                    MessageBox.Show("Data inserted!", "Done!");
            }
            catch (OleDbException e)
            {
                MessageBox.Show(e.ToString(), "Error");
            }
            finally
            {
                // Closing the connection to the database
                dbConnection.Close();
            }
        }

        public void Update()
        {
            OleDbConnection dbConnection = new OleDbConnection();

            try
            {
                // Creating a connection
                string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";
                dbConnection = new OleDbConnection(connectionString);

                // Executing a query to the database
                dbConnection.Open();

                // Updating all data
                bool err = false;
                foreach(Field f in fields)
                {
                    string query = "UPDATE table_name SET Name = '" + f.name + " WHERE id = " + f.id;
                    OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);

                    // Executing the request
                    if (dbCommand.ExecuteNonQuery() != 1)
                        err = true;
                }

               if(err)
                    MessageBox.Show("Updating error!", "Error!");
                else
                    MessageBox.Show("Data update!", "Done!");
            }
            catch (OleDbException e)
            {
                MessageBox.Show(e.ToString(), "Error");
            }
            finally
            {
                // Closing the connection to the database
                dbConnection.Close();
            }
        }

        public void Update(Field f)
        {
            OleDbConnection dbConnection = new OleDbConnection();

            try
            {
                // Creating a connection
                string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";
                dbConnection = new OleDbConnection(connectionString);

                // Executing a query to the database
                dbConnection.Open();
                string query = "UPDATE table_name SET Name = '" + f.name + " WHERE id = " + f.id;
                OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);

                // Executing the request
                if (dbCommand.ExecuteNonQuery() != 1)
                    MessageBox.Show("Updating error!", "Error!");
                else
                    MessageBox.Show("Data update!", "Done!");
            }
            catch (OleDbException e)
            {
                MessageBox.Show(e.ToString(), "Error");
            }
            finally
            {
                // Closing the connection to the database
                dbConnection.Close();
            }
        }

        public void Delete(Field f)
        {
            OleDbConnection dbConnection = new OleDbConnection();

            try
            {
                // Creating a connection
                string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";
                dbConnection = new OleDbConnection(connectionString);

                // Executing a query to the database
                dbConnection.Open();
                string query = "DELETE FROM table_name WHERE id = " + f.id;
                OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);

                // Executing the request
                if (dbCommand.ExecuteNonQuery() != 1)
                    MessageBox.Show("Deleting error!", "Error!");
                else
                {
                    MessageBox.Show("Data delete!", "Done!");
                    fields.Remove(f);
                }
            }
            catch (OleDbException e)
            {
                MessageBox.Show(e.ToString(), "Error");
            }
            finally
            {
                // Closing the connection to the database
                dbConnection.Close();
            }
        }

        public void Delete()
        {
            OleDbConnection dbConnection = new OleDbConnection();

            try
            {
                // Creating a connection
                string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";
                dbConnection = new OleDbConnection(connectionString);

                // Executing a query to the database
                dbConnection.Open();
                string query = "DELETE FROM table_name"; // clear table
                OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);

                // Executing the request
                if (dbCommand.ExecuteNonQuery() != 1)
                    MessageBox.Show("Deleting error!", "Error!");
                else
                {
                    MessageBox.Show("Data delete!", "Done!");
                }
            }
            catch (OleDbException e)
            {
                MessageBox.Show(e.ToString(), "Error");
            }
            finally
            {
                // Closing the connection to the database
                dbConnection.Close();
            }
        }
    }
}
