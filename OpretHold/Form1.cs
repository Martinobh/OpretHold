using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
//using Excel = Microsoft.Office;
//using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office;
//using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop;
using System.Data.OleDb;
using System.Web;





namespace OpretHold
{
    public partial class Form1 : Form
    {

        DataSet LokaleHoldFordeling = new DataSet();
        DataSet DBHold = new DataSet();
        
        DataSet Tnr = new DataSet();
        DataSet TraenerHold = new DataSet();
        

        int inc = 0;  

        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }


        private void firstTeam()
        {
            DataTable HoldTemp2 = DBHold.Tables[0];
            var result = HoldTemp2.AsEnumerable()
                                                .Where((row, index) => index == inc)
                                                .CopyToDataTable();

            DataSet SelectedData3 = new DataSet();
            SelectedData3.Tables.Add(result);

            DataRow DRrow = SelectedData3.Tables[0].Rows[0];

            DataTable HoldTemp3 = LokaleHoldFordeling.Tables[0];
            var result2 = HoldTemp3.AsEnumerable()
                                                .Where((row, index) => index == inc)
                                                .CopyToDataTable();

            DataSet SelectedData4 = new DataSet();
            SelectedData4.Tables.Add(result2);

            DataRow DRrow2 = SelectedData4.Tables[0].Rows[0];


     //        DataRow DRrow2 = SelectedData4.Tables[0].Rows[0];

            //--------------------
            //-- Her laver jeg træner tabellen

           // DataTable HoldTemp5 = LokaleHoldFordeling.Tables[0];
            /*       var result3 = HoldTemp4.AsEnumerable()
                                                       .Where((row, index) => index == inc)
                                                       .CopyToDataTable();
                   DataSet SelectedData5 = new DataSet();
                   SelectedData5.Tables.Add(result3);*/
         /*   HoldTemp5.Columns.Remove("trænernavn");
            HoldTemp5.Columns.Remove("niveau");
            HoldTemp5.Columns.Remove("gebyr");
            HoldTemp5.Columns.Remove("Ej stud");
            HoldTemp5.Columns.Remove("tnr");
            HoldTemp5.Columns.Remove("holdnavn");
            HoldTemp5.Columns.Remove("starttid");
            HoldTemp5.Columns.Remove("sluttid");
            HoldTemp5.Columns.Remove("startdato");
            HoldTemp5.Columns.Remove("slutdato");
            HoldTemp5.Columns.Remove("lokalekode");
            HoldTemp5.Columns.Remove("ugedag");
            HoldTemp5.Columns.Remove("periode");
            HoldTemp5.Columns.Remove("bemærkninger");
            dataGridView5.DataSource = HoldTemp5;*/
            //        DataRow DRrow2 = SelectedData4.Tables[0].Rows[0];




            //---------------------
            Holdkode.Text = DRrow["HoldKode"].ToString();
            Holdnavn.Text = DRrow2["Holdnavn"].ToString();
     //       DeltagerRefPris.Text = DRrow["DeltagerAfgReference"].ToString();
            DeltagerPris.Text = DRrow2["Gebyr"].ToString();
     //       EjStudRefPris.Text = DRrow["EjStudAfgReference"].ToString();
            EjStudPris.Text = DRrow2["Ej stud"].ToString();
     //       Bemaerkninger.Text = DRrow["Bemærkninger"].ToString();
            Koen.Text = DRrow["Køn"].ToString();
    //        Tid.Text = DRrow["Tid"].ToString();
            Niveau.Text = DRrow2["Niveau"].ToString();
    //        CB_Vis.Checked = Convert.ToBoolean(DRrow["Vis"].ToString());
    //        CB_Aktiv.Checked = Convert.ToBoolean(DRrow["Aktiv"].ToString());
            Holdpladser.Text = DRrow["Holdpladser"].ToString();
            ProcentIkkeStud.Text = DRrow["Procent ikke stud"].ToString();
    //        ExtraGebyr.Text = DRrow["ExtraGebyr"].ToString();
            Medlemsgebyr.Text = DRrow["medlemsgebyr"].ToString();
            AdminGebyr.Text = DRrow["admingebyr"].ToString();
    //        CB_Sommerhold.Checked = Convert.ToBoolean(DRrow["Sommerhold"].ToString());
   //         TraenerRefPris.Text = DRrow["TraenerPrisReference"].ToString();
            TraenerPris.Text = DRrow["trænerpris"].ToString();
            holdType.Text = DRrow["Holdtype"].ToString();
            holdtypeID.Text = DRrow["HoldtypeID"].ToString();
            Sportsgren.Text = DRrow["sportsgren"].ToString();
            Budgetteret.Text = DRrow["budgetteret"].ToString();
            VID.Text = DRrow["VID"].ToString();
            TraenerePerGang.Text = DRrow["trænere pr gang"].ToString();
            AntTraeninger.Text = DRrow["antal træninger"].ToString();
 /*           CB_TotalFriTimeindtastning.Checked = Convert.ToBoolean(DRrow["TotalFriTimeIndtastning"].ToString());
            CB_opkraevIkkeMedlemsgebyr.Checked = Convert.ToBoolean(DRrow["OpkrævIkkeMedlemsgebyr"].ToString());
            CB_Afmeldingsgebyr.Checked = Convert.ToBoolean(DRrow["afmeldegebyr"].ToString());
            CB_Flyttegebyr.Checked = Convert.ToBoolean(DRrow["flyttegebyr"].ToString());
            CB_FriTimeindtastning.Checked = Convert.ToBoolean(DRrow["FriTimeIndtastning"].ToString());
            CB_Venteliste.Checked = Convert.ToBoolean(DRrow["Venteliste"].ToString());
            CB_AabenFortidligereEjStud.Checked = Convert.ToBoolean(DRrow["Fase_aabenForTidligereEjStud"].ToString());
            CB_Fase2.Checked = Convert.ToBoolean(DRrow["Fase2"].ToString());
            CB_Fase1.Checked = Convert.ToBoolean(DRrow["Fase1"].ToString());
            CB_Klubtilmelding.Checked = Convert.ToBoolean(DRrow["Klubtilmelding"].ToString());
            CB_AabenTilmelding.Checked = Convert.ToBoolean(DRrow["AabenTilmelding"].ToString());
            CB_Parvistilmelding.Checked = Convert.ToBoolean(DRrow["ParvisTilmelding"].ToString());
            CB_Kursushold.Checked = Convert.ToBoolean(DRrow["KursusHold"].ToString());
  */
        }


        private string insertInDb()
        {

            System.Data.SqlClient.SqlConnection Connection1 = new System.Data.SqlClient.SqlConnection("server=sql.metalogic.dk; database=USGKontor;Persist Security Info=True;User ID=usgkontor;Password=1hihh1hihh");

            System.Data.SqlClient.SqlCommand cmd2 = new System.Data.SqlClient.SqlCommand();
            cmd2.CommandType = System.Data.CommandType.Text;
            string returnValueHID = "";
          /*    string stmt = "INSERT INTO USGkontor.Test4(T1, T2, T3) VALUES(@Holdkode, @Holdnavn, @TraenerPris)";

              SqlCommand cmd2 = new SqlCommand(stmt, Connection1);
              cmd2.Parameters.Add("@Holdkode", SqlDbType.VarChar, 100);
              cmd2.Parameters.Add("@Holdnavn", SqlDbType.VarChar, 100);
              cmd2.Parameters.Add("@TTraenerPris", SqlDbType.VarChar, 100);

              cmd2.Parameters["@Holdkode"].Value = Holdkode.Text;
              cmd2.Parameters["@Holdnavn"].Value = Holdnavn.Text;
              cmd2.Parameters["@TraenerPris"].Value = TraenerPris.Text;

              cmd2.ExecuteNonQuery();*/



              System.Data.SqlClient.SqlConnection sqlConnection1 = new System.Data.SqlClient.SqlConnection("server=sql.metalogic.dk; database=USGKontor;Persist Security Info=True;User ID=usgkontor;Password=1hihh1hihh");
        
              System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
              cmd.CommandType = System.Data.CommandType.Text;
              cmd.CommandText = "INSERT USGKontor.Test4 (T1, T2,T3)  VALUES(@Holdkode, @Holdnavn, @TraenerPris) ";
                                 
              cmd.Connection = sqlConnection1;

              cmd.Parameters.Add("@Holdkode", SqlDbType.VarChar, 100);
              cmd.Parameters.Add("@Holdnavn", SqlDbType.VarChar, 100);
              cmd.Parameters.Add("@TraenerPris", SqlDbType.VarChar, 100);

              cmd.Parameters["@Holdkode"].Value = Holdkode.Text;
              cmd.Parameters["@Holdnavn"].Value = Holdnavn.Text;
              cmd.Parameters["@TraenerPris"].Value = TraenerPris.Text;
              
            
              sqlConnection1.Open();
                 cmd.ExecuteNonQuery();


//-------------------
                 /*
                             string navn = "";
                             try
                             {

                                 SqlCommand command = sqlConnection1.CreateCommand();
                                 SqlTransaction transaction;

                                 transaction = sqlConnection1.BeginTransaction(IsolationLevel.ReadCommitted);

                                 command.Connection = sqlConnection1;
                                 command.Transaction = transaction;






                                 cmd.CommandText = "INSERT USGKontor.TrænerHold (Nr, Holdkode, HID)  VALUES(@TraenerNr, @Holdkode, @HID) ";

                                 cmd.Connection = sqlConnection1;

                                 cmd.Parameters.Add("@Nr", SqlDbType.VarChar, 100);
                                 cmd.Parameters.Add("@Holdkode", SqlDbType.VarChar, 100);
                                 cmd.Parameters.Add("@HID", SqlDbType.VarChar, 100);

                                 cmd.Parameters["@Nr"].Value = Holdkode.Text;
                                 cmd.Parameters["@Holdkode"].Value = Holdnavn.Text;
                                 cmd.Parameters["@HID"].Value = TraenerPris.Text;

                                 cmd.ExecuteNonQuery();



                                 cmd.CommandText = "INSERT USGKontor.TrænerHold (LokaleKode, Holdkode, Ugedag, Starttid, Sluttid, Startdato, Slutdato, HID, Periode, Fritraening)  VALUES(@LokaleKode, @Holdkode, @Ugedag, @Starttid, @Sluttid, @Startdato, @Slutdato, @HID, @Periode, @Fritraening) ";

                                 cmd.Connection = sqlConnection1;

                                 cmd.Parameters.Add("@LokaleKode", SqlDbType.VarChar, 100);
                                 cmd.Parameters.Add("@Holdkode", SqlDbType.VarChar, 100);
                                 cmd.Parameters.Add("@Ugedag", SqlDbType.VarChar, 100);
                                 cmd.Parameters.Add("@Starttid", SqlDbType.DateTime, 100);
                                 cmd.Parameters.Add("@Sluttid", SqlDbType.DateTime, 100);
                                 cmd.Parameters.Add("@Startdato", SqlDbType.Date, 100);
                                 cmd.Parameters.Add("@Slutdato", SqlDbType.Date, 100);
                                 cmd.Parameters.Add("@HID", SqlDbType.Int, 100);
                                 cmd.Parameters.Add("@Periode", SqlDbType.VarChar, 100);
                                 cmd.Parameters.Add("@Fritraening", SqlDbType.Bit, 100);

                                 cmd.Parameters["@LokaleKode"].Value = Holdkode.Text;
                                 cmd.Parameters["@Holdkode"].Value = Holdnavn.Text;
                                 cmd.Parameters["@Ugedag"].Value = Ugedag.Text;
                                 cmd.Parameters["@Starttid"].Value = Starttid.Text;
                                 cmd.Parameters["@Sluttid"].Value = Holdnavn.Text;
                                 cmd.Parameters["@Startdato"].Value = TraenerPris.Text;
                                 cmd.Parameters["@Slutdato"].Value = Holdkode.Text;
                                 cmd.Parameters["@HID"].Value = Holdnavn.Text;
                                 cmd.Parameters["@Periode"].Value = TraenerPris.Text;
                                 cmd.Parameters["@Fritraening"].Value = Holdkode.Text;


                                 cmd.ExecuteNonQuery();


                                 cmd3.CommandText = "INSERT USGKontor.Hold (Holdkode, Holdnavn, DeltagerAfgReference, Deltagerafg, EjStudAfgReference, EjStudAfg, Køn, Niveau, Vis, Aktiv, Holdplads, ProcentIkkeStud, Rate1, Rate2, ExtraGebyr, BetalingsFrist, MedlemsGebyr, AdminGebyr, SommerHold, TraenerPris, KursusHold, ParvisTilmelding, AabenTilmelding, KlubTilmelding, Fase1, Fase2, Fase_aabenForTidligereEjStud, Venteliste, HoldGruppe, sportsgrenid, Sportsgren, Ho1, StudenterPris, Budgetteret, VID, TrænerePerGang, AntTræninger, ventelistePladsGyldigAntalDage, FriTimeIndtastning, flyttegebyr, afmeldegebyr, opkrævIkkeMedlemsgebyr, TotalFriTimeIndtastning, PuljeLoen) " +
                                                                  " Output Inserted.HID VALUES(@Holdkode, @Holdnavn, @DeltagerAfgReference, @Deltagerafg, @EjStudAfgReference, @EjStudAfg, @Køn, @Niveau, @Vis, @Aktiv, @Holdplads, @ProcentIkkeStud, @Rate1, @Rate2, @ExtraGebyr, @BetalingsFrist, @MedlemsGebyr, @AdminGebyr, @SommerHold, @TraenerPris, @KursusHold, @ParvisTilmelding, @AabenTilmelding, @KlubTilmelding, @Fase1, @Fase2, @Fase_aabenForTidligereEjStud, @Venteliste, @HoldGruppe, @sportsgrenid, @Sportsgren, @Ho1, @StudenterPris, @Budgetteret, @VID, @TrænerePerGang, @AntTræninger, @ventelistePladsGyldigAntalDage, @FriTimeIndtastning, @flyttegebyr, @afmeldegebyr, @opkrævIkkeMedlemsgebyr, @TotalFriTimeIndtastning, @PuljeLoen) ";
                                 
                  returnValueHID = cmd3.ExecuteScalar().ToString();
                                
                  
                  cmd.Connection = sqlConnection1;




                                 cmd.Parameters.Add("@PuljeLoen", SqlDbType.Bit, 100);
                                 cmd.Parameters["@PuljeLoen"].Value = false;

                                 cmd.Parameters.Add("@TotalFriTimeIndtastning", SqlDbType.Bit, 100);
                                 cmd.Parameters["@TotalFriTimeIndtastning"].Value = false;

                                 cmd.Parameters.Add("@opkrævIkkeMedlemsgebyr", SqlDbType.Bit, 100);
                                 cmd.Parameters["@opkrævIkkeMedlemsgebyr"].Value = false;

                                 cmd.Parameters.Add("@afmeldegebyr", SqlDbType.Bit, 100);
                                 cmd.Parameters["@afmeldegebyr"].Value = false;

                                 cmd.Parameters.Add("@flyttegebyr", SqlDbType.Bit, 100);
                                 cmd.Parameters["@flyttegebyr"].Value = false;

                                 cmd.Parameters.Add("@FriTimeIndtastning", SqlDbType.Bit, 100);
                                 cmd.Parameters["@FriTimeIndtastning"].Value = false;

                                 cmd.Parameters.Add("@ventelistePladsGyldigAntalDage", SqlDbType.Int, 100);
                                 cmd.Parameters["@ventelistePladsGyldigAntalDage"].Value = 0;

                                 cmd.Parameters.Add("@AntTræninger", SqlDbType.Int, 100);
                                 cmd.Parameters["@AntTræninger"].Value = AntTraeninger.Text;


                                 cmd.Parameters.Add("@AntTræninger", SqlDbType.Int, 100);
                                 cmd.Parameters["@AntTræninger"].Value = AntTraeninger.Text;

                                 cmd.Parameters.Add("@TrænerePerGang", SqlDbType.Int, 100);
                                 cmd.Parameters["@TrænerePerGang"].Value = TraenerePerGang.Text;


                                 cmd.Parameters.Add("@VID", SqlDbType.Int, 100);
                                 cmd.Parameters["@VID"].Value = VID.Text;

                                 cmd.Parameters.Add("@Budgetteret", SqlDbType.Int, 100);
                                 cmd.Parameters["@Budgetteret"].Value = Budgetteret.Text;


                                 cmd.Parameters.Add("@StudenterPris", SqlDbType.Int, 100);
                                 cmd.Parameters["@StudenterPris"].Value = 0;

                                 cmd.Parameters.Add("@Ho1", SqlDbType.Bit, 100);
                                 cmd.Parameters["@Ho1"].Value = false;


                                 cmd.Parameters.Add("@Sportsgren", SqlDbType.VarChar, 100);
                                 cmd.Parameters["@Sportsgren"].Value = Sportsgren.Text;

                                 cmd.Parameters.Add("@sportsgrenid", SqlDbType.Int, 100);
                                 cmd.Parameters["@sportsgrenid"].Value =  // MANGLER SKAL HENTES VIA INSERT;      


                                 cmd.Parameters.Add("@HoldGruppe", SqlDbType.Bit, 100);
                                 cmd.Parameters["@HoldGruppe"].Value = false;


                                 cmd.Parameters.Add("@Venteliste", SqlDbType.Bit, 100);
                                 cmd.Parameters["@Venteliste"].Value = false;

                                 cmd.Parameters.Add("@Fase_aabenForTidligereEjStud", SqlDbType.Bit, 100);
                                 cmd.Parameters["@Fase_aabenForTidligereEjStud"].Value = false;

                                 cmd.Parameters.Add("@Fase2", SqlDbType.Bit, 100);
                                 cmd.Parameters["@Fase2"].Value = false;

                                 cmd.Parameters.Add("@Fase1", SqlDbType.Bit, 100);
                                 cmd.Parameters["@Fase1"].Value = false;
                                 cmd.Parameters.Add("@KlubTilmelding", SqlDbType.Bit, 100);
                                 cmd.Parameters["@KlubTilmelding"].Value = false;


                                 cmd.Parameters.Add("@AabenTilmelding", SqlDbType.Bit, 100);
                                 cmd.Parameters["@AabenTilmelding"].Value = false;

                                 cmd.Parameters.Add("@ParvisTilmelding", SqlDbType.Bit, 100);
                                 cmd.Parameters["@ParvisTilmelding"].Value = false;

                                 cmd.Parameters.Add("@KursusHold", SqlDbType.Bit, 100);
                                 cmd.Parameters["@KursusHold"].Value = false;

                                 cmd.Parameters.Add("@TraenerPris", SqlDbType.Int, 100);
                                 cmd.Parameters["@TraenerPris"].Value = TraenerPris.Text;

                                 cmd.Parameters.Add("@Holdkode", SqlDbType.VarChar, 100);
                                 cmd.Parameters.Add("@Holdnavn", SqlDbType.VarChar, 100);
                                 cmd.Parameters.Add("@DeltagerAfgReference", SqlDbType.Int, 100);
                                 cmd.Parameters.Add("@Deltagerafg", SqlDbType.Int, 100);
                                 cmd.Parameters.Add("@EjStudAfgReference", SqlDbType.Int, 100);
                                 cmd.Parameters.Add("@EjStudAfg", SqlDbType.Int, 100);
                                 cmd.Parameters.Add("@Køn", SqlDbType.VarChar, 100);
                                 cmd.Parameters.Add("@Niveau", SqlDbType.VarChar, 100);
                                 cmd.Parameters.Add("@Vis", SqlDbType.Bit, 100);

                                 cmd.Parameters.Add("@Aktiv", SqlDbType.Bit, 100);
                                 cmd.Parameters.Add("@Holdplads", SqlDbType.Int, 100);
                                 cmd.Parameters.Add("@ProcentIkkeStud", SqlDbType.Int, 100);
                                 cmd.Parameters.Add("@Rate1", SqlDbType.Int, 100);
                                 cmd.Parameters.Add("@Rate2", SqlDbType.Int, 100);
                                 cmd.Parameters.Add("@ExtraGebyr", SqlDbType.Int, 100);
                                 cmd.Parameters.Add("@MedlemsGebyr", SqlDbType.Int, 100);
                                 cmd.Parameters.Add("@AdminGebyr", SqlDbType.Int, 100);
                                 cmd.Parameters.Add("@SommerHold", SqlDbType.Bit, 100);


                                 cmd.Parameters["@Holdkode"].Value = Holdnavn.Text;
                                 cmd.Parameters["@Holdnavn"].Value = TraenerPris.Text;
                                 cmd.Parameters["@DeltagerAfgReference"].Value = DeltagerPris.Text; // DeltagerRefPris.Text;
                                 cmd.Parameters["@Deltagerafg"].Value = DeltagerPris.Text;
                                 cmd.Parameters["@EjStudAfgReference"].Value = EjStudPris.Text;//EjStudRefPris.Text;
                                 cmd.Parameters["@EjStudAfg"].Value = EjStudPris.Text;
                                 cmd.Parameters["@Køn"].Value = Koen.Text;
                                 cmd.Parameters["@Niveau"].Value = Niveau.Text;
                                 cmd.Parameters["@Vis"].Value = false;

                                 cmd.Parameters["@Aktiv"].Value = false;
                                 cmd.Parameters["@Holdplads"].Value = Holdpladser.Text;
                                 cmd.Parameters["@ProcentIkkeStud"].Value = ProcentIkkeStud.Text;
                                 cmd.Parameters["@Rate1"].Value = 0;
                                 cmd.Parameters["@Rate2"].Value = 0;
                                 cmd.Parameters["@ExtraGebyr"].Value = 0;
                                 cmd.Parameters["@MedlemsGebyr"].Value = Medlemsgebyr.Text;
                                 cmd.Parameters["@AdminGebyr"].Value = AdminGebyr.Text;
                                 cmd.Parameters["@SommerHold"].Value = false;


                                 cmd.Parameters.Add("@TraenerPrisReference", SqlDbType.VarChar, 100);
                                 cmd.Parameters["@TraenerPrisReference"].Value = TraenerPris.Text;






                                 transaction.Commit();
                             }
                             catch(Exception eee)
                             {
                
                
                             }

                             */


                 //      cmd.ExecuteNonQuery();
            
//---------------------------------------
            // Eksempel på hvordan jeg kan få data tilbage via sql

                 using (SqlCommand cmd3 = new SqlCommand("INSERT INTO USGKontor.Test4(T1,T2) output INSERTED.ID VALUES(@na,@occ)", sqlConnection1))
              {
                  cmd3.Parameters.AddWithValue("@na", Holdkode.Text);
                  cmd3.Parameters.AddWithValue("@occ", DeltagerPris.Text);
           //       cmd3.Parameters.AddWithValue("@occ", T3;
           //       sqlConnection1.Open();

                  returnValueHID = cmd3.ExecuteScalar().ToString();
             
           //       if (sqlConnection1.State == System.Data.ConnectionState.Open)
            //          sqlConnection1.Close();

                      
              }

//--------------------------------------
                return returnValueHID;
              Connection1.Close();
              sqlConnection1.Close();
        
        }





        private void button1_Click(object sender, EventArgs e)
        {

        //    insertInDb();

    /*        System.Data.SqlClient.SqlConnection sqlConnection1 = new System.Data.SqlClient.SqlConnection("server=sql.metalogic.dk; database=USGKontor;Persist Security Info=True;User ID=usgkontor;Password=1hihh1hihh");

            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
            cmd.CommandType = System.Data.CommandType.Text;
            cmd.CommandText = "INSERT USGKontor.Test4 (T1, T2,T3) VALUES ('" + Holdkode.Text + "', '" + DeltagerPris.Text + "','" + DeltagerRefPris.Text + "')";
            cmd.Connection = sqlConnection1;




            sqlConnection1.Open();
       //     cmd.ExecuteNonQuery();
            sqlConnection1.Close();
            */
            var HK = "";

              if (LokaleHoldFordeling != null)
                    {                                             

                            if (inc <= DBHold.Tables[0].Rows.Count)
                            {

                                DataTable HoldTemp2 = DBHold.Tables[0];
                                inc++; 
                                var result = HoldTemp2.AsEnumerable()
                                            .Where((row, index) => index == inc)
                                            .CopyToDataTable();
                                
                                DataSet SelectedData3 = new DataSet();                           
                                SelectedData3.Tables.Add(result);
                                DataRow DRrow = SelectedData3.Tables[0].Rows[0];

                                DataTable HoldTemp3 = LokaleHoldFordeling.Tables[0];
                                var result2 = HoldTemp3.AsEnumerable()
                                            .Where((row, index) => index == inc)
                                            .CopyToDataTable();

                                DataSet SelectedData4 = new DataSet();
                                SelectedData4.Tables.Add(result2);
                                DataRow DRrow2 = SelectedData4.Tables[0].Rows[0];


                                
                                Holdkode.Text = DRrow["HoldKode"].ToString();
                                Holdnavn.Text = DRrow2["Holdnavn"].ToString();
                             //   DeltagerRefPris.Text = DRrow["DeltagerAfgReference"].ToString();
                                DeltagerPris.Text = DRrow2["Gebyr"].ToString();
                             //   EjStudRefPris.Text = DRrow["EjStudAfgReference"].ToString();
                                EjStudPris.Text = DRrow2["Ej stud"].ToString();
                                Bemaerkninger.Text = DRrow2["Bemærkninger"].ToString();
                                Koen.Text = DRrow["Køn"].ToString();
                         //       Tid.Text = DRrow["Tid"].ToString();
                                Niveau.Text = DRrow2["Niveau"].ToString();
                        //        CB_Vis.Checked = Convert.ToBoolean(DRrow["Vis"].ToString());
                        //        CB_Aktiv.Checked = Convert.ToBoolean(DRrow["Aktiv"].ToString());
                                Holdpladser.Text = DRrow["Holdpladser"].ToString();
                                ProcentIkkeStud.Text = DRrow["Procent ikke stud"].ToString();
                        //        ExtraGebyr.Text = DRrow["ExtraGebyr"].ToString();
                                Medlemsgebyr.Text = DRrow["Medlemsgebyr"].ToString();
                                AdminGebyr.Text = DRrow["AdminGebyr"].ToString();
                         //       CB_Sommerhold.Checked = Convert.ToBoolean(DRrow["Sommerhold"].ToString());
                         //       TraenerRefPris.Text = DRrow["TraenerPrisReference"].ToString();
                                TraenerPris.Text = DRrow["TraenerPris"].ToString();
                                Sportsgren.Text = DRrow["sportsgren"].ToString();
                                Budgetteret.Text = DRrow["Budgetteret"].ToString();
                                VID.Text = DRrow["VID"].ToString();
                                TraenerePerGang.Text = DRrow["Trænere Pr Gang"].ToString();
                                AntTraeninger.Text = DRrow["Antal Træninger"].ToString();
                          //      CB_TotalFriTimeindtastning.Checked = Convert.ToBoolean(DRrow["TotalFriTimeIndtastning"].ToString());
                          //      CB_opkraevIkkeMedlemsgebyr.Checked = Convert.ToBoolean(DRrow["OpkrævIkkeMedlemsgebyr"].ToString());
                          //      CB_Afmeldingsgebyr.Checked = Convert.ToBoolean(DRrow["afmeldegebyr"].ToString());
                          //      CB_Flyttegebyr.Checked = Convert.ToBoolean(DRrow["flyttegebyr"].ToString());
                          //      CB_FriTimeindtastning.Checked = Convert.ToBoolean(DRrow["FriTimeIndtastning"].ToString());
                          //      CB_Venteliste.Checked = Convert.ToBoolean(DRrow["Venteliste"].ToString());
                          //      CB_AabenFortidligereEjStud.Checked = Convert.ToBoolean(DRrow["Fase_aabenForTidligereEjStud"].ToString());
                          //      CB_Fase2.Checked = Convert.ToBoolean(DRrow["Fase2"].ToString());
                          //      CB_Fase1.Checked = Convert.ToBoolean(DRrow["Fase1"].ToString());
                          //      CB_Klubtilmelding.Checked = Convert.ToBoolean(DRrow["Klubtilmelding"].ToString());
                          //      CB_AabenTilmelding.Checked = Convert.ToBoolean(DRrow["AabenTilmelding"].ToString());
                          //      CB_Parvistilmelding.Checked = Convert.ToBoolean(DRrow["ParvisTilmelding"].ToString());
                          //      CB_Kursushold.Checked = Convert.ToBoolean(DRrow["KursusHold"].ToString());
                                holdType.Text = DRrow["Holdtype"].ToString();
                                holdtypeID.Text = DRrow["Holdtypeid"].ToString();
                                dataGridView3.DataSource = SelectedData3.Tables[0];
                              
                         













                            }   
                    }
              
              updateGrids();



           /* var connectionString = ConfigurationManager.ConnectionStrings["CharityManagement"].ConnectionString;
            SqlConnection myConnection = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["UsgDatabase"].ConnectionString);
            myConnection.Open();

            string stmt = "INSERT INTO USGkontor.Test4(T1, T2, T3) VALUES(@Holdkode, @T2, @T3)";

            SqlCommand cmd2 = new SqlCommand(stmt, myConnection);
            cmd.Parameters.Add("@Holdkode", SqlDbType.VarChar, 100);
            cmd.Parameters.Add("@T2", SqlDbType.VarChar, 100);
            cmd.Parameters.Add("@T3", SqlDbType.VarChar, 100);

            cmd.Parameters["@Holdkode"].Value = textBox1.Text;
            cmd.Parameters["@T2"].Value = textBox2.Text;
            cmd.Parameters["@T3"].Value = textBox2.Text;

            cmd2.ExecuteNonQuery();


            myConnection.Close();
          */

        //    var lines = File.ReadLines("input.txt").Skip(1);

         /*   using (System.IO.StreamReader reader = new System.IO.StreamReader(fileStream))
            {
                // Read the first line from the file and write it the textbox.
                string temp = reader.ReadLine();
                string[] words = temp.Split(' ');
                textBox1.Text = words[0];
                textBox2.Text = words[1];
                textBox3.Text = words[2];
                textBox4.Text = words[3];
                textBox5.Text = words[4];

                foreach (string word in words)
                {
                    Console.WriteLine(word);
                }
                //     textBox1.Text = reader.ReadLine();
            }*/
            
    //        using (var db = new dev.Models.usg.USGKontorEntities())
        }

        private void updateGrids() 
        {


         //   DataTable table = LokaleHoldFordeling.Tables[0];
            DataTable tbl1 = new DataTable();
            DataTable tbl = LokaleHoldFordeling.Tables[0];
            DataRow[] dr = tbl.Select("Holdkode = '" + Holdkode.Text + "'");
            if (dr.Length > 0)
            {
                tbl1 = dr.CopyToDataTable();
            }
            DataSet SelectedData = new DataSet();
            SelectedData.Tables.Add(tbl1);
            dataGridView1.DataSource = SelectedData.Tables[0];



            DataTable TabelTilTraenereLocal = new DataTable();
            DataTable tblA = TabelTilTraenereGlobal;//.Tables[0];
            DataRow[] dr2 = tblA.Select("Holdkode = '" + Holdkode.Text + "'");
            if (dr2.Length > 0)
            {
                TabelTilTraenereLocal = dr2.CopyToDataTable();
            }
            DataSet SelectedData2 = new DataSet();
            SelectedData2.Tables.Add(TabelTilTraenereLocal);
            dataGridView5.DataSource = SelectedData2.Tables[0];




            DataTable TabelTilLokaleHoldFordeling = new DataTable();
            DataTable tblB = LokaleHoldFordelingGlobal;//.Tables[0];
            DataRow[] dr3 = tblB.Select("Holdkode = '" + Holdkode.Text + "'");
            if (dr3.Length > 0)
            {
                TabelTilLokaleHoldFordeling = dr3.CopyToDataTable();
            }
            DataSet SelectedData3 = new DataSet();
            SelectedData3.Tables.Add(TabelTilLokaleHoldFordeling);
            dataGridView4.DataSource = SelectedData3.Tables[0];




            /*
            DataTable table2 = Tnr.Tables[0];
            DataTable TabelTilTraenere = new DataTable();
            DataTable tbl3 = Tnr.Tables[0];
            DataRow[] dr2 = tbl3.Select("Holdkode = '" + Holdkode.Text + "'");
            if (dr2.Length > 0)
            {
                TabelTilTraenere = dr2.CopyToDataTable();
            }
 
            DataSet SelectedData2 = new DataSet();
            SelectedData2.Tables.Add(TabelTilTraenere);
            dataGridView2.DataSource = SelectedData2.Tables[0];
        */
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }
        System.IO.Stream fileStream;

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        DataTable LokaleHoldFordelingGlobal = new DataTable();
        DataTable TabelTilTraenereGlobal = new DataTable();
        private BindingSource bindingSource1 = new BindingSource();
        DataTable TabelTilTraenere = new DataTable();
        DataTable tbl4 = new DataTable();
        void openfile()
        {

            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;

            string str;
            int rCnt = 0;
            int cCnt = 0;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
         //   xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\Martin\Desktop\Martin_Juli_2015\Arbejde\sommer 2014\IMPORT_ALLE_HOLD.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Holdimport\HoldTilImport.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;     
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            range = xlWorkSheet.UsedRange;
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            
        //    string Connstr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:\\Users\\Martin\\Desktop\\Martin_Juli_2015\\Arbejde\\sommer 2014\\IMPORT_ALLE_HOLD.xlsx; " + 
        //                     "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';";

            string Connstr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:\\Holdimport\\HoldTilImport.xlsx; " +
                         "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';";
            OleDbConnection conn = new OleDbConnection(Connstr);

            //------Hold Start-----------------
         //   string DBHoldSQL = "SELECT * FROM [DBHold$]";
            string DBHoldSQL = "SELECT Holdkode, Holdpladser, Køn, Medlemsgebyr, trænerpris, budgetteret, [trænere pr gang], [antal træninger], [Procent ikke stud], admingebyr, sportsgren, VID, Holdtype, HoldtypeID FROM [Ark2$]";

            OleDbCommand DBHoldcmd = new OleDbCommand(DBHoldSQL, conn);
            OleDbDataAdapter DBHoldDa = new OleDbDataAdapter(DBHoldcmd);
            DBHoldDa.Fill(DBHold);

            DataTable HoldTemp2 = DBHold.Tables[0];
            var result = HoldTemp2.AsEnumerable()
                        .Where((row, index) => index == 0)
                        .CopyToDataTable();

            DataSet SelectedData3 = new DataSet();
            SelectedData3.Tables.Add(result);
            dataGridView3.DataSource = SelectedData3.Tables[0];
            Holdkode.Text = SelectedData3.Tables[0].Rows[0]["HoldKode"].ToString();
            //------Hold Slut-----------------


           
     //       string strSQL = "SELECT * FROM [LokaleHoldFordeling$]";
            string strSQL = "SELECT HoldKode, Trænernavn, LokaleKode, Ugedag, Starttid, Sluttid, Startdato, Slutdato, Holdnavn, Niveau, Periode, Gebyr, [Ej stud], Bemærkninger, Tnr FROM [Ark1$]";
            // 'trænere pr. gang', 'Antal træninger', 'Procent ikke stud',

            OleDbCommand cmd = new OleDbCommand(strSQL, conn);           
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(LokaleHoldFordeling);
        //    dataGridView1.DataSource = LokaleHoldFordeling.Tables[0];
           

            DataTable table = LokaleHoldFordeling.Tables[0];

            DataTable tbl1 = new DataTable();
            DataTable tbl = LokaleHoldFordeling.Tables[0];
            DataRow[] dr = tbl.Select("Holdkode = '" + Holdkode.Text + "'");
            if (dr.Length > 0)
            {
                tbl1 = dr.CopyToDataTable();
            }

            DataSet SelectedData = new DataSet();
            SelectedData.Tables.Add(tbl1);
            dataGridView1.DataSource = SelectedData.Tables[0];
            dataGridView1.Columns[0].DefaultCellStyle = new DataGridViewCellStyle { Format = "hh:mm:ss" };

          //  dataGridView1.Rows[5].Cells[3].Value = Convert.ToDateTime("hh:mm:ss");
            //----------------------------------------------------

            DataTable HoldTemp5 = LokaleHoldFordeling.Tables[0];

            TabelTilTraenereGlobal = LokaleHoldFordeling.Tables[0].Copy();
            TabelTilTraenereGlobal.Columns.Remove("niveau");
            TabelTilTraenereGlobal.Columns.Remove("gebyr");
            TabelTilTraenereGlobal.Columns.Remove("Ej stud");
            //  TabelTilTraenere.Columns.Remove("tnr");
            TabelTilTraenereGlobal.Columns.Remove("holdnavn");
            TabelTilTraenereGlobal.Columns.Remove("starttid");
            TabelTilTraenereGlobal.Columns.Remove("sluttid");
            TabelTilTraenereGlobal.Columns.Remove("startdato");
            TabelTilTraenereGlobal.Columns.Remove("slutdato");
            TabelTilTraenereGlobal.Columns.Remove("lokalekode");
            TabelTilTraenereGlobal.Columns.Remove("ugedag");
            TabelTilTraenereGlobal.Columns.Remove("periode");
            TabelTilTraenereGlobal.Columns.Remove("bemærkninger");







            DataTable tbl3 = LokaleHoldFordeling.Tables[0].Copy();
            DataRow[] dr2 = tbl3.Select("Holdkode = '" + Holdkode.Text + "'");
            if (dr2.Length > 0)
            {
                TabelTilTraenere = dr2.CopyToDataTable();
            }
            /*       var result3 = HoldTemp4.AsEnumerable()
                                                       .Where((row, index) => index == inc)
                                                       .CopyToDataTable();
                   DataSet SelectedData5 = new DataSet();
                   SelectedData5.Tables.Add(result3);*/
          //  TabelTilTraenere.Columns.Remove("trænernavn");
            TabelTilTraenere.Columns.Remove("niveau");
            TabelTilTraenere.Columns.Remove("gebyr");
            TabelTilTraenere.Columns.Remove("Ej stud");
          //  TabelTilTraenere.Columns.Remove("tnr");
            TabelTilTraenere.Columns.Remove("holdnavn");
            TabelTilTraenere.Columns.Remove("starttid");
            TabelTilTraenere.Columns.Remove("sluttid");
            TabelTilTraenere.Columns.Remove("startdato");
            TabelTilTraenere.Columns.Remove("slutdato");
            TabelTilTraenere.Columns.Remove("lokalekode");
            TabelTilTraenere.Columns.Remove("ugedag");
            TabelTilTraenere.Columns.Remove("periode");
            TabelTilTraenere.Columns.Remove("bemærkninger");
               dataGridView5.DataSource = TabelTilTraenere;
            //        DataRow DRrow2 = SelectedData4.Tables[0].Rows[0];

//----------------------------------------


               //Her fylder jeg data ind i tabellen til lokaleholdfordeling
               DataTable HoldTemp4 = LokaleHoldFordeling.Tables[0].Copy();

               
               DataTable tbl5 = LokaleHoldFordeling.Tables[0].Copy();

               DataRow[] dr4 = tbl5.Select("Holdkode = '" + Holdkode.Text + "'");
               if (dr4.Length > 0)
               {
                   tbl4 = dr2.CopyToDataTable();
               }

            /*       var result3 = HoldTemp4.AsEnumerable()
                                                          .Where((row, index) => index == inc)
                                                          .CopyToDataTable();
                      DataSet SelectedData5 = new DataSet();
                      SelectedData5.Tables.Add(result3);*/

               LokaleHoldFordelingGlobal = LokaleHoldFordeling.Tables[0].Copy();
               tbl4.Columns.Remove("trænernavn");
               tbl4.Columns.Remove("niveau");
               tbl4.Columns.Remove("gebyr");
               tbl4.Columns.Remove("Ej stud");
               tbl4.Columns.Remove("tnr");
               tbl4.Columns.Remove("holdnavn");
               tbl4.Columns.Add("VisSomNr", typeof(System.Int32));
               tbl4.Columns.Add("Timefaktor", typeof(System.Int32));
               tbl4.Columns.Add("Fritraening");

               LokaleHoldFordelingGlobal.Columns.Remove("trænernavn");
               LokaleHoldFordelingGlobal.Columns.Remove("niveau");
               LokaleHoldFordelingGlobal.Columns.Remove("gebyr");
               LokaleHoldFordelingGlobal.Columns.Remove("Ej stud");
               LokaleHoldFordelingGlobal.Columns.Remove("tnr");
               LokaleHoldFordelingGlobal.Columns.Remove("holdnavn");
               LokaleHoldFordelingGlobal.Columns.Add("VisSomNr", typeof(System.Int32));
               LokaleHoldFordelingGlobal.Columns.Add("Timefaktor", typeof(System.Int32));
               LokaleHoldFordelingGlobal.Columns.Add("Fritraening");

               dataGridView4.DataSource = tbl4;


            /*
            string TraenerHoldSQL = "SELECT * FROM [TrænerHold$]";
          

            OleDbCommand Traenercmd = new OleDbCommand(TraenerHoldSQL, conn);
            OleDbDataAdapter Traenerda = new OleDbDataAdapter(Traenercmd);
            Traenerda.Fill(TraenerHold);

          //  string Connstr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:\\Users\\Martin\\Desktop\\Martin_Juli_2015\\Arbejde\\sommer 2014\\IMPORT_ALLE_HOLD.xlsx; " +
         //        "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';";


            string traenerSQL = "SELECT * FROM [TrænerHold$]";
            OleDbCommand Trncmd = new OleDbCommand(traenerSQL, conn);
            OleDbDataAdapter daTr = new OleDbDataAdapter(Trncmd);
            daTr.Fill(Tnr);


            DataTable table2 = Tnr.Tables[0];
            DataTable tbl3 = Tnr.Tables[0];
            DataTable TabelTilTraenere = new DataTable();
            DataRow[] dr2 = tbl3.Select("Holdkode = '" + Holdkode.Text + "'");
            if (dr2.Length > 0)
            {
                TabelTilTraenere = dr2.CopyToDataTable();
            }
            DataSet SelectedData2 = new DataSet();
            SelectedData2.Tables.Add(TabelTilTraenere);
            dataGridView2.DataSource = SelectedData2.Tables[0];
*/
            firstTeam();

          //  updateGrids();

            


            

     //       Connstr.Close();

     //       "Provider=Microsoft.ACE.OLEDB.12.0: Data Source=" + path +   


        /*    string mySheet = @"C:\Users\Martin\Desktop\Martin_Juli_2015\Arbejde\sommer 2014\Holdtabel.xlsx";
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbooks books = excelApp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook sheet = books.Open(mySheet);

             Microsoft.Office.Interop.Excel.Worksheet worksheet = sheet.Sheets.get_Item(1); // (Excel.Worksheet)sheets.get_Item(1);
                for (int i = 1; i <= 10; i++)
                {
                    Microsoft.Office.Interop.Excel.Range range = worksheet.get_Range("A" + i.ToString(), "J" + i.ToString());
                    System.Array myvalues = (System.Array)range.Cells.Value;
               //     textBox2.Text = range.Cells.Value.toString();
                  //  string[] strArray = ConvertToStringArray(myvalues);
                }
                books.Close();*/
            conn.Close();
        }





        private void button2_Click(object sender, EventArgs e)
        {
            openfile();
            /*

            int size = -1;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Browse Text Files";

            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;

            openFileDialog1.DefaultExt = "txt";
            openFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
     //       openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|Early Excel files (*.xls)|*.xls";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.ShowReadOnly = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
               // textBox1.Text = openFileDialog1.FileName;
                fileStream = openFileDialog1.OpenFile();
                using (System.IO.StreamReader reader = new System.IO.StreamReader(fileStream))
                {
                    // Read the first line from the file and write it the textbox.
                    string temp = reader.ReadLine();
                    string[] words = temp.Split(' ');
                    textBox1.Text = words[0];
                    textBox2.Text = words[1];
                    textBox3.Text = words[2];
                    textBox4.Text = words[3];
                    textBox5.Text = words[4];
                    
                 //   foreach (string word in words)
                 //   {
                 //       Console.WriteLine(word);
                //    }
               //     textBox1.Text = reader.ReadLine();
                }
           //     fileStream.Close();

            */

        /*    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object file = path;
            object nullobj = System.Reflection.Missing.Value;

            wordApp.Documents.doc = wordApp.Documents.Open(ref file, ref nullobj, ref nullobj,
                                                   ref nullobj, ref nullobj, ref nullobj,
                                                   ref nullobj, ref nullobj, ref nullobj,
                                                   ref nullobj, ref nullobj, ref nullobj);

            string result = doc.Content.Text.Trim();
            doc.Close();
            return result;*/

       //      OpenFileDialog openFileDialog1 = new OpenFileDialog();

       /*     openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Browse Text Files";

            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;

            openFileDialog1.DefaultExt = "txt";
     //       openFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|Early Excel files (*.xls)|*.xls";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.ShowReadOnly = true;
            */


            Microsoft.Office.Interop.Excel.Application wordApp = new Microsoft.Office.Interop.Excel.Application();

           // Microsoft.Office.Interop.Excel.Workbook theWorkbook = Microsoft.Office.Interop.Excel.Workbooks.Open(openFileDialog1.FileName, 0, true, 5,"", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);
     //       Microsoft.Office.Interop.Excel.Application.Workbooks.Open(@"C:\Test\YourWorkbook.xlsx");
      //      Microsoft.Office.Interop.Excel.Sheets sheets = theWorkbook.Worksheets;
       /*     Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(1);
            for (int i = 1; i <= 10; i++)
            {
                Microsoft.Office.Interop.Excel.Range range = worksheet.get_Range("A" + i.ToString(), "J" + i.ToString());
                System.Array myvalues = (System.Array)range.Cells.Value;
            //    string[] strArray = ConvertToStringArray(myvalues);
            }*/

       //     Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)this.Sheets["Sheet2"];
       //     sheet.Select(Type.Missing);

            

          //  Microsoft.Office.Interop.Excel.Workbook wb = ThisApplication.Workbooks.Add(Type.Missing);
        /*    Microsoft.Office.Interop.Excel.Workbook wb = ThisApplication.Workbooks.Open( 
    "C:\\YourPath\\Yourworkbook.xls", 
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, 
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, 
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, 
    Type.Missing, Type.Missing);
            */

      //      this.openFileDialog1.FileName = "*.xls";
      //      if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
     //       {
      /*          Excel.Workbook theWorkbook = ExcelObj.Workbooks.Open(
                   openFileDialog1.FileName, 0, true, 5,
                    "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false,
                    0, true);
                Excel.Sheets sheets = theWorkbook.Worksheets;
                Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);
                for (int i = 1; i <= 10; i++)
                {
                    Excel.Range range = worksheet.get_Range("A" + i.ToString(), "J" + i.ToString());
                    System.Array myvalues = (System.Array)range.Cells.Value;
                    string[] strArray = ConvertToStringArray(myvalues);
                }*/
    //        }


               //string[] lines = System.IO.File.ReadAllLines(@"openFileDialog1.FileName");

            // Display the file contents by using a foreach loop.

            //  System.Console.WriteLine("Contents of WriteLines2.txt = " + lines[0]);

           //   foreach (string line in lines)
            //  { }


            }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        
       
        }
    }

