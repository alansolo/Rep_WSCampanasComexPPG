using EntidadesCampanasPPG.BD;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UtilidadesCampanasPPG;

namespace DatosCampanasPPG
{
    public class BD
    {
        private SqlConnection sqlConnection;
        private SqlCommand command;
        protected string connectionString;

        public BD()
        {
            connectionString = ConfigurationManager.ConnectionStrings["MktCampanas"].ToString();
        }
        public BD(string dataBaseConnection)
        {
            connectionString = ConfigurationManager.ConnectionStrings[dataBaseConnection].ToString();
        }
        protected string ExecuteScalar(string spName, List<SqlParameter> parameters)
        {
            string result = string.Empty;

            using (sqlConnection = new SqlConnection(connectionString))
            {
                try
                {
                    command = new SqlCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = spName;
                    foreach (SqlParameter item in parameters)
                        command.Parameters.Add(item);

                    command.Connection = sqlConnection;
                    sqlConnection.Open();
                    result = (string)command.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    //Write log error
                    ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": Method: "+ ex.Source + " Error: " + ex.Message);
                    throw ex;
                }
            }

            return result;
        }
        protected int ExecuteNonQuery(string spName, List<SqlParameter> parameters)
        {
            int result = 0;

            using (sqlConnection = new SqlConnection(connectionString))
            {
                try
                {
                    command = new SqlCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = spName;
                    foreach (SqlParameter item in parameters)
                        command.Parameters.Add(item);

                    command.Connection = sqlConnection;
                    sqlConnection.Open();
                    result = command.ExecuteNonQuery();

                }
                catch (Exception ex)
                {
                    //Write log error
                    ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + spName + ": Method: " + ex.Source + " Error: " + ex.Message);
                    throw ex;
                }
            }

            return result;
        }
        protected int ExecuteNonQueryBitacora(string spName, List<SqlParameter> parameters)
        {
            int result = 0;

            using (sqlConnection = new SqlConnection(connectionString))
            {
                try
                {
                    command = new SqlCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = spName;
                    foreach (SqlParameter item in parameters)
                        command.Parameters.Add(item);

                    command.Connection = sqlConnection;
                    sqlConnection.Open();
                    result = command.ExecuteNonQuery();


                }
                catch (Exception ex)
                {
                    //Write log error
                    ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + spName + ": Method: " + ex.Source + " Error: " + ex.Message);
                    throw ex;
                }
            }

            return result;
        }
        //protected int ExecuteNonQuery(string spName, List<SqlParameterGroup> parametersGroup)
        //{
        //    SqlTransaction transaction;
        //    int result = 0;

        //    using (sqlConnection = new SqlConnection(connectionString))
        //    {
        //        sqlConnection.Open();
        //        transaction = sqlConnection.BeginTransaction(IsolationLevel.ReadCommitted);

        //        foreach (SqlParameterGroup paramGroup in parametersGroup)
        //        {
        //            try
        //            {
        //                command = new SqlCommand();
        //                command.Connection = sqlConnection;
        //                command.Transaction = transaction;
        //                command.CommandType = CommandType.StoredProcedure;
        //                command.CommandText = spName;

        //                foreach (SqlParameter item in paramGroup.Parameter)
        //                    command.Parameters.Add(item);

        //                result = command.ExecuteNonQuery();
        //            }
        //            catch (Exception ex)
        //            {
        //                result = 0;
        //                transaction.Rollback();

        //                //Write log error
        //                LogFile.WriteLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": Method: " + ex.Source + " Error: " + ex.Message);
        //                throw ex;
        //            }
        //        }

        //        transaction.Commit();
        //    }

        //    return result;
        //}
        protected DataTable ExecuteNonQueryDataTable(string spName, List<SqlParameter> parameters)
        {
            DataSet dataSet = null;
            DataTable dataTable = new DataTable();

            using (sqlConnection = new SqlConnection(connectionString))
            {
                try
                {
                    command = new SqlCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = spName;

                    foreach (SqlParameter item in parameters)
                        command.Parameters.Add(item);

                    command.Connection = sqlConnection;
                    sqlConnection.Open();

                    SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
                    dataSet = new DataSet("dsResult");
                    dataAdapter.Fill(dataSet);

                    if (dataSet.Tables.Count > 0)
                    {
                        dataTable = dataSet.Tables[0];
                    }
                }
                catch (Exception ex)
                {
                    //Write log error
                    ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": Method: " + ex.Source + " Error: " + ex.Message);
                    throw ex;
                }
            }

            return dataTable;
        }
        
        //protected DataTable ExecuteNonQueryDataTable(string spName, List<SqlParameterGroup> parametersGroup)
        //{
        //    DataSet dataSet = null;
        //    DataTable dataTable = new DataTable();
        //    DataTable dtTotal = new DataTable();
        //    SqlTransaction transaction;

        //    using (sqlConnection = new SqlConnection(connectionString))
        //    {
        //        sqlConnection.Open();
        //        transaction = sqlConnection.BeginTransaction(IsolationLevel.ReadCommitted);

        //        foreach (SqlParameterGroup paramGroup in parametersGroup)
        //        {
        //            try
        //            {
        //                command = new SqlCommand();
        //                command.Connection = sqlConnection;
        //                command.Transaction = transaction;
        //                command.CommandType = CommandType.StoredProcedure;
        //                command.CommandText = spName;

        //                foreach (SqlParameter item in paramGroup.Parameter)
        //                    command.Parameters.Add(item);                      

        //                SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
        //                dataSet = new DataSet("dsResult");
        //                dataAdapter.Fill(dataSet);

        //                if (dataSet.Tables.Count > 0)
        //                {
        //                    dataTable = dataSet.Tables[0];

        //                    dtTotal.Merge(dataTable);
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                transaction.Rollback();

        //                //Write log error
        //                LogFile.WriteLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": Method: " + ex.Source + " Error: " + ex.Message);
        //                throw ex;
        //            }
        //        }

        //        transaction.Commit();
        //    }

        //    return dtTotal;
        //}
        protected DataTable ExecuteDataTable(string spName, List<SqlParameter> parameters)
        {
            DataSet dataSet = null;
            DataTable dataTable = new DataTable();

            using (sqlConnection = new SqlConnection(connectionString))
            {
                try
                {
                    command = new SqlCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = spName;
                    command.CommandTimeout = 120;

                    foreach (SqlParameter item in parameters)
                        command.Parameters.Add(item);

                    command.Connection = sqlConnection;
                    sqlConnection.Open();

                    SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
                    dataSet = new DataSet("dsResult");
                    dataAdapter.Fill(dataSet);

                    if (dataSet.Tables.Count > 0)
                    {
                        dataTable = dataSet.Tables[0];
                    }
                }
                catch (Exception ex)
                {
                    //Write log error
                    ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": Method: " + ex.Source + " Error: " + ex.Message);
                    throw ex;
                }
            }

            return dataTable;
        }
        protected DataSet ExecuteDataSet(string spName, List<SqlParameter> parameters)
        {
            DataSet dataSet = null;

            using (sqlConnection = new SqlConnection(connectionString))
            {
                try
                {
                    command = new SqlCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = spName;
                    //CINCO MINUTOS TIEMPO DE ESPERA PARA EJECUTAR RENTABILIDAD
                    command.CommandTimeout = 300;
                    foreach (SqlParameter item in parameters)
                        command.Parameters.Add(item);

                    command.Connection = sqlConnection;
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
                    dataSet = new DataSet("dsResult");
                    dataAdapter.Fill(dataSet);
                }
                catch (Exception ex)
                {
                    //Write log error
                    ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": Method: " + ex.Source + " Error: " + ex.Message);
                    throw ex;
                }
            }

            return dataSet;
        }

        protected int ExecuteNonQueryCampana(string SpNameCampana, List<SqlParameter> ParametersCampana, 
                                            string SpNameUpdateCampana, List<SqlParameter> ParametersUpdateCampana,
                                            string SpNamePublicidad, List<SqlParameterGroup> ListParametersGroupPublicidad,
                                            string SpNameEliminarPublicidad, List<SqlParameter> ParameterEliminarPublicidad,
                                            string SpNameKroma, List<SqlParameterGroup> ListParametersGroupKroma,
                                            string SpNameEliminarKroma, List<SqlParameter> ParametersEliminarKroma,
                                            string SpNameSignoVital, List<SqlParameterGroup> ListParameterGroupSignoVital,
                                            string SpNameEliminarSignoVital, List<SqlParameter> ParameterEliminarSignoVital)
        {
            int result = 0;
            int idCampana = 0;
            SqlParameter sqlParameterIdCampanaOut = new SqlParameter();
            SqlParameter sqlParameterIdCampanaIn = new SqlParameter();
            SqlParameter sqlParameterResult = new SqlParameter();
            SqlTransaction transaction;

            using (sqlConnection = new SqlConnection(connectionString))
            {
                sqlConnection.Open();
                transaction = sqlConnection.BeginTransaction(IsolationLevel.ReadCommitted);

                try
                {
                    command = new SqlCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = SpNameCampana;
                    command.Transaction = transaction;
                    foreach (SqlParameter item in ParametersCampana)
                        command.Parameters.Add(item);

                    command.Connection = sqlConnection;
                    result = command.ExecuteNonQuery();

                    if(result > 0 || result == -1)
                    {
                        sqlParameterIdCampanaOut = ParametersCampana.Where(n => n.ParameterName == "@ID_Campc").FirstOrDefault();

                        if (sqlParameterIdCampanaOut != null && sqlParameterIdCampanaOut.Value != null)
                        {
                            idCampana = Convert.ToInt32(sqlParameterIdCampanaOut.Value);

                            //if (result < 0)
                            //{
                            //    transaction.Rollback();
                            //    result = -1;

                            //    ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + ParametersEliminarKroma + "Error: No se elimino correctamente la informacion de productos en la base de datos.");

                            //    throw new Exception("Error: No se guardo o actualizo correctamente la informacion en la base de datos.");
                            //}

                            //ACTUALIZAR DATOS CAMPANA
                            command.Parameters.Clear();

                            command.CommandText = SpNameUpdateCampana;

                            sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                            sqlParameterIdCampanaIn.Value = idCampana;

                            command.Parameters.Add(sqlParameterIdCampanaIn);

                            foreach (SqlParameter item in ParametersUpdateCampana)
                                command.Parameters.Add(item);

                            command.Connection = sqlConnection;
                            result = command.ExecuteNonQuery();


                            //ELIMINAR PUBLICIDAD ANTERIOR
                            command.Parameters.Clear();

                            command.CommandText = SpNameEliminarPublicidad;

                            sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                            sqlParameterIdCampanaIn.Value = idCampana;

                            command.Parameters.Add(sqlParameterIdCampanaIn);

                            foreach (SqlParameter item in ParameterEliminarPublicidad)
                                command.Parameters.Add(item);

                            command.Connection = sqlConnection;
                            result = command.ExecuteNonQuery();

                            foreach (SqlParameterGroup paramGroup in ListParametersGroupPublicidad)
                            {
                                command.Parameters.Clear();

                                command.CommandText = SpNamePublicidad;

                                sqlParameterIdCampanaIn.ParameterName = "@ID_Campaña";
                                sqlParameterIdCampanaIn.Value = idCampana;

                                command.Parameters.Add(sqlParameterIdCampanaIn);
                                
                                foreach (SqlParameter item in paramGroup.ListSqlParameter)
                                    command.Parameters.Add(item);

                                command.Connection = sqlConnection;
                                result = command.ExecuteNonQuery();

                                if(result <= 0)
                                {
                                    transaction.Rollback();
                                    result = -1;

                                    ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNamePublicidad + "Error: No se guardo o actualizo correctamente la informacion en la base de datos.");

                                    throw new Exception("Error: No se guardo o actualizo correctamente la informacion en la base de datos.");
                                }
                            }


                            //ELIMINAR PARTICIPACION KROMA ANTERIOR
                            command.Parameters.Clear();

                            command.CommandText = SpNameEliminarKroma;

                            sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                            sqlParameterIdCampanaIn.Value = idCampana;

                            command.Parameters.Add(sqlParameterIdCampanaIn);

                            foreach (SqlParameter item in ParametersEliminarKroma)
                                command.Parameters.Add(item);

                            command.Connection = sqlConnection;
                            result = command.ExecuteNonQuery();

                            foreach (SqlParameterGroup paramGroup in ListParametersGroupKroma)
                            {
                                command.Parameters.Clear();

                                command.CommandText = SpNameKroma;

                                sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                sqlParameterIdCampanaIn.Value = idCampana;

                                command.Parameters.Add(sqlParameterIdCampanaIn);

                                foreach (SqlParameter item in paramGroup.ListSqlParameter)
                                    command.Parameters.Add(item);

                                command.Connection = sqlConnection;
                                result = command.ExecuteNonQuery();

                                if (result <= 0)
                                {
                                    transaction.Rollback();
                                    result = -1;

                                    ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:hh:mm:ss") + ": SP:" + SpNameKroma + "Error: No se guardo o actualizo correctamente la informacion en la base de datos.");

                                    throw new Exception("Error: No se guardo o actualizo correctamente la informacion en la base de datos.");

                                }
                            }


                            //ELIMINAR SIGNO VITAL
                            command.Parameters.Clear();

                            command.CommandText = SpNameEliminarSignoVital;

                            sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                            sqlParameterIdCampanaIn.Value = idCampana;

                            command.Parameters.Add(sqlParameterIdCampanaIn);

                            foreach (SqlParameter item in ParameterEliminarSignoVital)
                                command.Parameters.Add(item);

                            command.Connection = sqlConnection;
                            result = command.ExecuteNonQuery();

                            foreach (SqlParameterGroup paramGroup in ListParameterGroupSignoVital)
                            {
                                command.Parameters.Clear();

                                command.CommandText = SpNameSignoVital;

                                sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                sqlParameterIdCampanaIn.Value = idCampana;

                                command.Parameters.Add(sqlParameterIdCampanaIn);

                                foreach (SqlParameter item in paramGroup.ListSqlParameter)
                                    command.Parameters.Add(item);

                                command.Connection = sqlConnection;
                                result = command.ExecuteNonQuery();

                                if (result <= 0)
                                {
                                    transaction.Rollback();
                                    result = -1;

                                    ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:hh:mm:ss") + ": SP:" + SpNameSignoVital + "Error: No se guardo o actualizo correctamente la informacion en la base de datos.");

                                    throw new Exception("Error: No se guardo o actualizo correctamente la informacion en la base de datos.");

                                }
                            }


                            transaction.Commit();
                            result = 1;

                        }
                        else
                        {
                            transaction.Rollback();

                            result = -1;

                            ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:hh:mm:ss") + ": SP:" + SpNameCampana + "Error: No se guardo o actualizo correctamente la informacion de la campaña en la base de datos.");

                            throw new Exception("Error: No se guardo o actualizo correctamente la informacion de la campaña en la base de datos.");
                        }
                    }
                    else
                    {
                        transaction.Rollback();

                        result = -1;

                        ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNameCampana + ":Error: No se encontro el parametro @ID_Campc, para el identificador de la campaña");

                        throw new Exception("Error: No se encontro el parametro @ID_Campc, para el identificador de la campaña");

                    }

                }
                catch (Exception ex)
                {
                    transaction.Rollback();

                    result = -1;

                    //Write log error
                    ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNameCampana + ": Method: " + ex.Source + " Error: " + ex.Message);
                    throw ex;
                }
            }

            return result;
        }
        protected int ExecuteNonQueryCampana(string SpNameCampana, List<SqlParameter> ParametersCampana, 
                                            string SpNameCronograma, List<SqlParameterGroup> ListParameterGroupCronograma,
                                            string SpNameRegionEliminar, List<SqlParameter> ParametersRegionEliminar,
                                            string SpNameRegion, List<SqlParameterGroup> ListParameterGroupRegion,
                                            string SpNameSubCanalEliminar, List<SqlParameter> ParametersSubCanalEliminar,
                                            string SpNameSubCanal, List<SqlParameterGroup> ListParameterGroupSubCanal
                                            )
        {
            int result = 0;
            int idCampana = 0;
            SqlParameter sqlParameterIdCampanaOut = new SqlParameter();
            SqlParameter sqlParameterIdCampanaIn = new SqlParameter();
            SqlParameter sqlParameterResult = new SqlParameter();
            SqlTransaction transaction;

            using (sqlConnection = new SqlConnection(connectionString))
            {
                sqlConnection.Open();
                transaction = sqlConnection.BeginTransaction(IsolationLevel.ReadCommitted);

                try
                {
                    command = new SqlCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = SpNameCampana;
                    command.Transaction = transaction;
                    foreach (SqlParameter item in ParametersCampana)
                        command.Parameters.Add(item);

                    command.Connection = sqlConnection;
                    result = command.ExecuteNonQuery();

                    if (result > 0)
                    {
                        sqlParameterIdCampanaOut = ParametersCampana.Where(n => n.ParameterName == "@ID_Campc").FirstOrDefault();

                        if (sqlParameterIdCampanaOut != null && sqlParameterIdCampanaOut.Value != null)
                        {
                            idCampana = Convert.ToInt32(sqlParameterIdCampanaOut.Value);

                            //AGREGAR ACTIVIDADES CRONOGRAMA
                            foreach (SqlParameterGroup paramGroup in ListParameterGroupCronograma)
                            {
                                command.Parameters.Clear();

                                command.CommandText = SpNameCronograma;

                                sqlParameterIdCampanaIn.ParameterName = "@IDCampania";
                                sqlParameterIdCampanaIn.Value = idCampana;

                                command.Parameters.Add(sqlParameterIdCampanaIn);

                                foreach (SqlParameter item in paramGroup.ListSqlParameter)
                                    command.Parameters.Add(item);

                                command.Connection = sqlConnection;
                                result = command.ExecuteNonQuery();

                                //if (result <= -1)
                                //{
                                //    transaction.Rollback();
                                //    result = -1;

                                //    ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNameCronograma + "Error: No se guardo o actualizo correctamente la informacion de cronograma en la base de datos.");

                                //    throw new Exception("Error: No se guardo o actualizo correctamente la informacion de cronograma en la base de datos.");
                                //}
                            }

                            //ELIMINAR REGIONES ANTERIORES
                            command.Parameters.Clear();

                            command.CommandText = SpNameRegionEliminar;

                            sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                            sqlParameterIdCampanaIn.Value = idCampana;

                            command.Parameters.Add(sqlParameterIdCampanaIn);

                            command.Transaction = transaction;
                            foreach (SqlParameter item in ParametersRegionEliminar)
                                command.Parameters.Add(item);

                            command.Connection = sqlConnection;
                            result = command.ExecuteNonQuery();


                            //AGREGAR REGION CAMPAÑA
                            foreach (SqlParameterGroup paramGroup in ListParameterGroupRegion)
                            {
                                command.Parameters.Clear();

                                command.CommandText = SpNameRegion;

                                sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                sqlParameterIdCampanaIn.Value = idCampana;

                                command.Parameters.Add(sqlParameterIdCampanaIn);

                                foreach (SqlParameter item in paramGroup.ListSqlParameter)
                                    command.Parameters.Add(item);

                                command.Connection = sqlConnection;
                                result = command.ExecuteNonQuery();
                            }


                            //ELIMINAR SUB CANALES ANTERIORES
                            command.Parameters.Clear();

                            command.CommandText = SpNameSubCanalEliminar;

                            sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                            sqlParameterIdCampanaIn.Value = idCampana;

                            command.Parameters.Add(sqlParameterIdCampanaIn);

                            command.Transaction = transaction;
                            foreach (SqlParameter item in ParametersRegionEliminar)
                                command.Parameters.Add(item);

                            command.Connection = sqlConnection;
                            result = command.ExecuteNonQuery();


                            //AGREGAR SUB CANAL CAMPAÑA
                            foreach (SqlParameterGroup paramGroup in ListParameterGroupSubCanal)
                            {
                                command.Parameters.Clear();

                                command.CommandText = SpNameSubCanal;

                                sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                sqlParameterIdCampanaIn.Value = idCampana;

                                command.Parameters.Add(sqlParameterIdCampanaIn);

                                foreach (SqlParameter item in paramGroup.ListSqlParameter)
                                    command.Parameters.Add(item);

                                command.Connection = sqlConnection;
                                result = command.ExecuteNonQuery();
                            }



                            transaction.Commit();
                            result = 1;
                        }
                        else
                        {
                            transaction.Rollback();

                            result = -1;

                            ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:hh:mm:ss") + ": SP:" + SpNameCampana + "Error: No se guardo o actualizo correctamente la informacion de la campaña en la base de datos.");

                            throw new Exception("Error: No se guardo o actualizo correctamente la informacion de la campaña en la base de datos.");
                        }
                    }
                    else
                    {
                        transaction.Rollback();

                        result = -1;

                        ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNameCampana + ":Error: No se encontro el parametro @ID_Campc, para el identificador de la campaña");

                        throw new Exception("Error: No se encontro el parametro @ID_Campc, para el identificador de la campaña");

                    }

                }
                catch (Exception ex)
                {
                    transaction.Rollback();

                    result = -1;

                    //Write log error
                    ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNameCampana + ": Method: " + ex.Source + " Error: " + ex.Message);
                    throw ex;
                }
            }

            return result;
        }
        protected DataTable ExecuteNonQueryProductoCampana(string SpNameCampana, List<SqlParameter> ParametersCampana,
                                                    string SpNameRegalo, List<SqlParameterGroup> ListParametersGroupRegalo,
                                                    string SPNameEliminarRegalo, List<SqlParameter> ParametersEliminarRegalo,
                                                    string SpNameMultiplo, List<SqlParameterGroup> ListParametersGroupMultiplo,
                                                    string SPNameEliminarMultiplo, List<SqlParameter> ParametersEliminarMultiplo,
                                                    string SpNameDescuento, List<SqlParameterGroup> ListParametersGroupDescuento,
                                                    string SPNameEliminarDescuento, List<SqlParameter> ParametersEliminarDescuento,
                                                    string SpNameVolumen, List<SqlParameterGroup> ListParametersGroupVolumen,
                                                    string SPNameEliminarVolumen, List<SqlParameter> ParametersEliminarVolumen,
                                                    string SpNameKit, List<SqlParameterGroup> ListParametersGroupKit,
                                                    string SPNameEliminarKit, List<SqlParameter> ParametersEliminarKit,
                                                    string SpNameCombo, List<SqlParameterGroup> ListParametersGroupCombo,
                                                    string SPNameEliminarCombo, List<SqlParameter> ParametersEliminarCombo,
                                                    string SpNameTienda, List<SqlParameterGroup> ListParametersGroupTienda,
                                                    string SPNameEliminarTienda, List<SqlParameter> ParametersEliminarTienda,
                                                    string SpNameTiendaExclusion, List<SqlParameterGroup> ListParametersGroupTiendaExclusion,
                                                    string SPNameEliminarTiendaExclusion, List<SqlParameter> ParametersEliminarTiendaExclusion,
                                                    string SPNameAlcance, List<SqlParameterGroup> ListParametersGroupAlcance,
                                                    string SPNameEliminarAlcance, List<SqlParameter> ParametersEliminarAlcance,
                                                    string SPNameMostrarLinea, List<SqlParameter> ParametersMostrarLinea)//,
                                                    //DataTable dtTienda)
        {
            int result = 0;
            int idCampana = 0;

            SqlParameter sqlParameterIdCampanaOut = new SqlParameter();
            SqlParameter sqlParameterIdCampanaIn = new SqlParameter();
            SqlParameter sqlParameterResult = new SqlParameter();
            SqlTransaction transaction;

            DataSet dataSet = null;
            DataTable dataTable = null;

            using (sqlConnection = new SqlConnection(connectionString))
            {
                sqlConnection.Open();

                using (transaction = sqlConnection.BeginTransaction(IsolationLevel.ReadCommitted))
                {

                    try
                    {
                        command = new SqlCommand();
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandText = SpNameCampana;
                        command.Transaction = transaction;
                        //Tiempo 10 minutos
                        command.CommandTimeout = 600;

                        foreach (SqlParameter item in ParametersCampana)
                            command.Parameters.Add(item);

                        command.Connection = sqlConnection;
                        result = command.ExecuteNonQuery();

                        if (result > 0 || result == -1)
                        {
                            sqlParameterIdCampanaOut = ParametersCampana.Where(n => n.ParameterName == "@ID_Campc").FirstOrDefault();

                            if (sqlParameterIdCampanaOut != null && sqlParameterIdCampanaOut.Value != null && sqlParameterIdCampanaOut.Value.ToString() != string.Empty)
                            {
                                idCampana = Convert.ToInt32(sqlParameterIdCampanaOut.Value);

                                #region REGALO

                                command.Parameters.Clear();

                                command.CommandText = SPNameEliminarRegalo;

                                sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                sqlParameterIdCampanaIn.Value = idCampana;

                                command.Parameters.Add(sqlParameterIdCampanaIn);

                                foreach (SqlParameter item in ParametersEliminarRegalo)
                                    command.Parameters.Add(item);

                                command.Connection = sqlConnection;
                                result = command.ExecuteNonQuery();

                                foreach (SqlParameterGroup paramGroup in ListParametersGroupRegalo)
                                {
                                    command.Parameters.Clear();

                                    command.CommandText = SpNameRegalo;

                                    sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                    sqlParameterIdCampanaIn.Value = idCampana;

                                    command.Parameters.Add(sqlParameterIdCampanaIn);

                                    foreach (SqlParameter item in paramGroup.ListSqlParameter)
                                        command.Parameters.Add(item);

                                    command.Connection = sqlConnection;
                                    result = command.ExecuteNonQuery();

                                    if (result <= 0)
                                    {
                                        //transaction.Rollback();
                                        //result = -1;

                                        //ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNameRegalo + "Error: No se guardo o actualizo correctamente la informacion en la base de datos.");

                                        throw new Exception("Error: No se guardo o actualizo correctamente la informacion en la base de datos.");
                                    }
                                }

                                #endregion

                                #region MULTIPLO

                                command.Parameters.Clear();

                                command.CommandText = SPNameEliminarMultiplo;

                                sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                sqlParameterIdCampanaIn.Value = idCampana;

                                command.Parameters.Add(sqlParameterIdCampanaIn);

                                foreach (SqlParameter item in ParametersEliminarMultiplo)
                                    command.Parameters.Add(item);

                                command.Connection = sqlConnection;
                                result = command.ExecuteNonQuery();

                                foreach (SqlParameterGroup paramGroup in ListParametersGroupMultiplo)
                                {
                                    command.Parameters.Clear();

                                    command.CommandText = SpNameMultiplo;

                                    sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                    sqlParameterIdCampanaIn.Value = idCampana;

                                    command.Parameters.Add(sqlParameterIdCampanaIn);

                                    foreach (SqlParameter item in paramGroup.ListSqlParameter)
                                        command.Parameters.Add(item);

                                    command.Connection = sqlConnection;
                                    result = command.ExecuteNonQuery();

                                    if (result <= 0)
                                    {
                                        //transaction.Rollback();
                                        //result = -1;

                                        //ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNameMultiplo + "Error: No se guardo o actualizo correctamente la informacion en la base de datos.");

                                        throw new Exception("Error: No se guardo o actualizo correctamente la informacion en la base de datos.");
                                    }
                                }

                                #endregion

                                #region DESCUENTO

                                command.Parameters.Clear();

                                command.CommandText = SPNameEliminarDescuento;

                                sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                sqlParameterIdCampanaIn.Value = idCampana;

                                command.Parameters.Add(sqlParameterIdCampanaIn);

                                foreach (SqlParameter item in ParametersEliminarDescuento)
                                    command.Parameters.Add(item);

                                command.Connection = sqlConnection;
                                result = command.ExecuteNonQuery();

                                foreach (SqlParameterGroup paramGroup in ListParametersGroupDescuento)
                                {
                                    command.Parameters.Clear();

                                    command.CommandText = SpNameDescuento;

                                    sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                    sqlParameterIdCampanaIn.Value = idCampana;

                                    command.Parameters.Add(sqlParameterIdCampanaIn);

                                    foreach (SqlParameter item in paramGroup.ListSqlParameter)
                                        command.Parameters.Add(item);

                                    command.Connection = sqlConnection;
                                    result = command.ExecuteNonQuery();

                                    if (result <= 0)
                                    {
                                        //transaction.Rollback();
                                        //result = -1;

                                        //ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNameDescuento + "Error: No se guardo o actualizo correctamente la informacion en la base de datos.");

                                        throw new Exception("Error: No se guardo o actualizo correctamente la informacion en la base de datos.");
                                    }
                                }

                                #endregion

                                #region VOLUMEN

                                command.Parameters.Clear();

                                command.CommandText = SPNameEliminarVolumen;

                                sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                sqlParameterIdCampanaIn.Value = idCampana;

                                command.Parameters.Add(sqlParameterIdCampanaIn);

                                foreach (SqlParameter item in ParametersEliminarVolumen)
                                    command.Parameters.Add(item);

                                command.Connection = sqlConnection;
                                result = command.ExecuteNonQuery();

                                foreach (SqlParameterGroup paramGroup in ListParametersGroupVolumen)
                                {
                                    command.Parameters.Clear();

                                    command.CommandText = SpNameVolumen;

                                    sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                    sqlParameterIdCampanaIn.Value = idCampana;

                                    command.Parameters.Add(sqlParameterIdCampanaIn);

                                    foreach (SqlParameter item in paramGroup.ListSqlParameter)
                                        command.Parameters.Add(item);

                                    command.Connection = sqlConnection;
                                    result = command.ExecuteNonQuery();

                                    if (result <= 0)
                                    {
                                        //transaction.Rollback();
                                        //result = -1;

                                        //ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNameVolumen + "Error: No se guardo o actualizo correctamente la informacion en la base de datos.");

                                        throw new Exception("Error: No se guardo o actualizo correctamente la informacion en la base de datos.");
                                    }
                                }

                                #endregion

                                #region KIT

                                command.Parameters.Clear();

                                command.CommandText = SPNameEliminarKit;

                                sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                sqlParameterIdCampanaIn.Value = idCampana;

                                command.Parameters.Add(sqlParameterIdCampanaIn);

                                foreach (SqlParameter item in ParametersEliminarKit)
                                    command.Parameters.Add(item);

                                command.Connection = sqlConnection;
                                result = command.ExecuteNonQuery();

                                foreach (SqlParameterGroup paramGroup in ListParametersGroupKit)
                                {
                                    command.Parameters.Clear();

                                    command.CommandText = SpNameKit;

                                    sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                    sqlParameterIdCampanaIn.Value = idCampana;

                                    command.Parameters.Add(sqlParameterIdCampanaIn);

                                    foreach (SqlParameter item in paramGroup.ListSqlParameter)
                                        command.Parameters.Add(item);

                                    command.Connection = sqlConnection;
                                    result = command.ExecuteNonQuery();

                                    if (result <= 0)
                                    {
                                        //transaction.Rollback();
                                        //result = -1;

                                        //ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNameKit + "Error: No se guardo o actualizo correctamente la informacion en la base de datos.");

                                        throw new Exception("Error: No se guardo o actualizo correctamente la informacion en la base de datos.");
                                    }
                                }

                                #endregion

                                #region COMBO

                                command.Parameters.Clear();

                                command.CommandText = SPNameEliminarCombo;

                                sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                sqlParameterIdCampanaIn.Value = idCampana;

                                command.Parameters.Add(sqlParameterIdCampanaIn);

                                foreach (SqlParameter item in ParametersEliminarCombo)
                                    command.Parameters.Add(item);

                                command.Connection = sqlConnection;
                                result = command.ExecuteNonQuery();

                                foreach (SqlParameterGroup paramGroup in ListParametersGroupCombo)
                                {
                                    command.Parameters.Clear();

                                    command.CommandText = SpNameCombo;

                                    sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                    sqlParameterIdCampanaIn.Value = idCampana;

                                    command.Parameters.Add(sqlParameterIdCampanaIn);

                                    foreach (SqlParameter item in paramGroup.ListSqlParameter)
                                        command.Parameters.Add(item);

                                    command.Connection = sqlConnection;
                                    result = command.ExecuteNonQuery();

                                    if (result <= 0)
                                    {
                                    //    transaction.Rollback();
                                    //    result = -1;

                                    //    ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNameCombo + "Error: No se guardo o actualizo correctamente la informacion en la base de datos.");

                                        throw new Exception("Error: No se guardo o actualizo correctamente la informacion en la base de datos.");
                                    }
                                }

                                #endregion

                                //dtTienda.AsEnumerable().ToList().ForEach(n =>
                                //{
                                //    n["ID_Campania"] = idCampana;
                                //});

                                //using (var sqlBulk = new SqlBulkCopy(sqlConnection, SqlBulkCopyOptions.KeepIdentity, transaction))
                                //{
                                //    sqlBulk.DestinationTableName = "CampaniaTiendas";
                                //    sqlBulk.WriteToServer(dtTienda);
                                //}


                                #region CLIENTES

                                command.Parameters.Clear();

                                command.CommandText = SPNameEliminarTienda;

                                sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                sqlParameterIdCampanaIn.Value = idCampana;

                                command.Parameters.Add(sqlParameterIdCampanaIn);

                                foreach (SqlParameter item in ParametersEliminarTienda)
                                    command.Parameters.Add(item);

                                command.Connection = sqlConnection;
                                result = command.ExecuteNonQuery();

                                foreach (SqlParameterGroup paramGroup in ListParametersGroupTienda)
                                {
                                    command.Parameters.Clear();

                                    command.CommandText = SpNameTienda;

                                    sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                    sqlParameterIdCampanaIn.Value = idCampana;

                                    command.Parameters.Add(sqlParameterIdCampanaIn);

                                    foreach (SqlParameter item in paramGroup.ListSqlParameter)
                                        command.Parameters.Add(item);

                                    command.Connection = sqlConnection;
                                    result = command.ExecuteNonQuery();
                                    

                                    if (result <= 0 && result != -1)
                                    {
                                        //transaction.Rollback();
                                        //result = -1;

                                        //ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNameTienda + "Error: No se guardo o actualizo correctamente la informacion en la base de datos.");

                                        throw new Exception("Error: No se guardo o actualizo correctamente la informacion en la base de datos.");
                                    }
                                }

                                #endregion

                                #region CLIENTES EXCLUSION

                                //command.Parameters.Clear();

                                //command.CommandText = SPNameEliminarTiendaExclusion;

                                //sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                //sqlParameterIdCampanaIn.Value = idCampana;

                                //command.Parameters.Add(sqlParameterIdCampanaIn);

                                //foreach (SqlParameter item in ParametersEliminarTiendaExclusion)
                                //    command.Parameters.Add(item);

                                //command.Connection = sqlConnection;
                                //result = command.ExecuteNonQuery();

                                foreach (SqlParameterGroup paramGroup in ListParametersGroupTiendaExclusion)
                                {
                                    command.Parameters.Clear();

                                    command.CommandText = SpNameTienda;

                                    sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                    sqlParameterIdCampanaIn.Value = idCampana;

                                    command.Parameters.Add(sqlParameterIdCampanaIn);

                                    foreach (SqlParameter item in paramGroup.ListSqlParameter)
                                        command.Parameters.Add(item);

                                    command.Connection = sqlConnection;
                                    result = command.ExecuteNonQuery();


                                    if (result <= 0 && result != -1)
                                    {
                                        //transaction.Rollback();
                                        //result = -1;

                                        //ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNameTienda + "Error: No se guardo o actualizo correctamente la informacion en la base de datos.");

                                        throw new Exception("Error: No se guardo o actualizo correctamente la informacion en la base de datos.");
                                    }
                                }

                                #endregion

                                #region ALCANCE

                                command.Parameters.Clear();

                                command.CommandText = SPNameEliminarAlcance;

                                sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                sqlParameterIdCampanaIn.Value = idCampana;

                                command.Parameters.Add(sqlParameterIdCampanaIn);

                                foreach (SqlParameter item in ParametersEliminarAlcance)
                                    command.Parameters.Add(item);

                                command.Connection = sqlConnection;
                                result = command.ExecuteNonQuery();

                                foreach (SqlParameterGroup paramGroup in ListParametersGroupAlcance)
                                {
                                    command.Parameters.Clear();

                                    command.CommandText = SPNameAlcance;

                                    sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                                    sqlParameterIdCampanaIn.Value = idCampana;

                                    command.Parameters.Add(sqlParameterIdCampanaIn);

                                    foreach (SqlParameter item in paramGroup.ListSqlParameter)
                                        command.Parameters.Add(item);

                                    command.Connection = sqlConnection;
                                    result = command.ExecuteNonQuery();

                                    if (result <= 0)
                                    {
                                        //transaction.Rollback();
                                        //result = -1;

                                        //ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SPNameAlcance + "Error: No se guardo o actualizo correctamente la informacion en la base de datos.");

                                        throw new Exception("Error: No se guardo o actualizo correctamente la informacion en la base de datos.");
                                    }
                                }

                                #endregion

                            }
                            else
                            {
                                //transaction.Rollback();
                                //transaction.Dispose();

                                //result = -1;

                                //ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNameCampana + ":Error: No se encontro el parametro @ID_Campc, para el identificador de la campaña");

                                throw new Exception("Error: No se encontro la Clave de Campaña, revise que la informacion de la Campaña sea correcta e intente de nuevo.");
                            }


                            transaction.Commit();
                            transaction.Dispose();
                            result = 1;

                            #region MostrarLinea

                            command.Parameters.Clear();

                            command.CommandText = SPNameMostrarLinea;

                            sqlParameterIdCampanaIn.ParameterName = "@ID_Campania";
                            sqlParameterIdCampanaIn.Value = idCampana;

                            command.Parameters.Add(sqlParameterIdCampanaIn);

                            foreach (SqlParameter item in ParametersMostrarLinea)
                                command.Parameters.Add(item);

                            //CINCO MINUTOS TIEMPO DE ESPERA PARA EJECUTAR RENTABILIDAD
                            command.CommandTimeout = 600;

                            command.Connection = sqlConnection;

                            //sqlConnection.Open();

                            SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
                            dataSet = new DataSet("dsResult");
                            dataAdapter.Fill(dataSet);

                            if (dataSet.Tables.Count > 0)
                            {
                                dataTable = dataSet.Tables[0];
                            }

                            #endregion
                        }
                        else
                        {
                            //transaction.Rollback();
                            //transaction.Dispose();

                            //result = -1;

                            //ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:hh:mm:ss") + ": SP:" + SpNameCampana + "Error: No se guardo o actualizo correctamente la informacion de la campaña en la base de datos.");

                            throw new Exception("Error: No se guardo o actualizo correctamente la informacion de la campaña en la base de datos.");
                        }

                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        transaction.Dispose();

                        result = -1;

                        //Write log error
                        ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpNameCampana + ": Method: " + ex.Source + " Error: " + ex.Message);
                        throw ex;
                    }
                }
            }

            return dataTable;
        }
        protected int ExecuteNonQueryTrans(string SpName, List<SqlParameterGroup> ListParametersGroup)
        {
            int result = 0;
            SqlTransaction transaction;

            using (sqlConnection = new SqlConnection(connectionString))
            {
                sqlConnection.Open();
                transaction = sqlConnection.BeginTransaction(IsolationLevel.ReadCommitted);

                try
                {
                    foreach (SqlParameterGroup sqlParameterGroup in ListParametersGroup)
                    {
                        command = new SqlCommand();
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandText = SpName;
                        command.Transaction = transaction;

                        foreach (SqlParameter item in sqlParameterGroup.ListSqlParameter)
                            command.Parameters.Add(item);

                        command.Connection = sqlConnection;
                        result = command.ExecuteNonQuery();
                    }

                    transaction.Commit();

                }
                catch (Exception ex)
                {
                    transaction.Rollback();

                    //Write log error
                    ArchivoLog.EscribirLog(null, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + ": SP:" + SpName + ": Method: " + ex.Source + " Error: " + ex.Message);
                    throw ex;
                }
            }

            return result;
        }

    }
}
