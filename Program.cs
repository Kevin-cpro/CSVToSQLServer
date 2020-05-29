using System;
using System.Xml;
using System.IO;
using System.Data.SqlClient;
using System.Linq;
using System.Collections.Generic;
using System.Security.Cryptography;

namespace CSVToSQLServer
{
    class Program
    {
        struct Colonnes
        {
            public string Nom;
            public string CP;
            public string Ville;
            public string Email;

            public Colonnes(string Nom, string CP, string Ville, string Email)
            {
                this.Nom = Nom;
                this.CP = CP;
                this.Ville = Ville;
                this.Email = Email;
            }
        }
        private static void UpdateDatas(SqlConnection connection, string queryString)
        {
                SqlCommand command = new SqlCommand(queryString, connection);
                command.ExecuteNonQuery();
        }
        private static List<Colonnes> SelectDatas(SqlConnection connection, string queryString)
        {
            List<Colonnes> output = new List<Colonnes>();

            SqlCommand command = new SqlCommand(
                    queryString, connection);
            using (SqlDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    output.Add(new Colonnes(reader[0].ToString(), reader[1].ToString(), reader[2].ToString(),reader[3].ToString()));
                }
            }
            return output;
        }

        static void Main(string[] args)
        {
            var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser("Test.csv");
            parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
            parser.SetDelimiters(new string[] { ";" });

            SqlConnection connection = new SqlConnection("Server = SRV-THEREFORE\\THEREFORE; Database = Therefore; Trusted_Connection = True;");
            connection.Open();

            using (StreamWriter file = new StreamWriter(@"log.csv"))
            {
                while (!parser.EndOfData)
                {
                    string[] row = parser.ReadFields();

                    if (row[6].Contains('\''))
                    {
                        row[6] = row[6].Replace('\'', ' ');
                    }
                    if (row[0].Contains('\''))
                    {
                        row[0] = row[0].Replace("\'", "\'\'");
                    }

                    Colonnes fromCsv = new Colonnes(row[0], row[5], row[6], row[9]);
                    string Statut = "";
                    string Commentaire = "";

                    if (fromCsv.Email != "")
                    {
                        if (!fromCsv.Email.Contains("@"))
                        {
                            Statut = "BDD non modifiee";
                            Commentaire = "Le format d'email est incoherent (manque l'@ ?)";
                        }
                        else
                        {
                            List<Colonnes> listFromDatabase = SelectDatas(connection, "SELECT Nom, CP, Ville, Email FROM dbo.TheCat10 WHERE Nom = '" + fromCsv.Nom + "' AND CP = '" + fromCsv.CP + "' AND Ville = '" + fromCsv.Ville + "';");
                            if (listFromDatabase.Count == 0)
                            {
                                Statut = "BDD non modifiee";
                                Commentaire = "La BDD ne contient aucun correspondant au Nom, CP, Ville";
                            }
                            else
                            {
                                if (listFromDatabase.Count > 1)
                                {
                                    Statut = "BDD non modifiee";
                                    Commentaire = "La BDD contient plus d'un enregistrement correspondant au Nom, CP, Ville";
                                }
                                else
                                {
                                    Colonnes fromDB = listFromDatabase[0];
                                    if (fromDB.Email.Contains('@'))
                                    {
                                        Statut = "BDD non modifiee";
                                        Commentaire = "La BDD contient un enregistrement avec un email qui semble valide (" + fromDB.Email + ")";
                                    }
                                    else
                                    {
                                        try
                                        {
                                            UpdateDatas(connection, "UPDATE TheCat10 SET Email = '" + fromCsv.Email + "' WHERE Nom = '" + fromCsv.Nom + "' AND CP = '" + fromCsv.CP + "' AND Ville = '" + fromCsv.Ville + "';");
                                            //Console.WriteLine("UPDATE TheCat10 SET Email = '" + fromCsv.Email + "' WHERE Nom = '" + fromCsv.Nom + "' AND CP = '" + fromCsv.CP + "' AND Ville = '" + fromCsv.Ville + "';");
                                            Statut = "BDD modifiee";
                                        }
                                        catch (SqlException se) {
                                            Statut = "BDD non modifiee";
                                            Commentaire = "Exception lors de l'update de la base: " + se.Message;
                                        }
                                    }
                                    
                                }

                            }
                        }
                    }
                    else
                    {
                        Statut = "BDD non modifiee";
                        Commentaire = "Pas d'email pour l'enregistrement";
                    }
                    file.WriteLine(fromCsv.Nom + ";" + fromCsv.CP + ";" + fromCsv.Ville + ";" + fromCsv.Email + ";" + Statut + ";" + Commentaire);
                }
            }
            connection.Close();
        }
    }
}
