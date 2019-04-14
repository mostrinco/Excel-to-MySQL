using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;


namespace Excel_a_Mysql_JCMM
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogoAbrirArchivo.ShowDialog();
            txtDireccionArchivo.Text = DialogoAbrirArchivo.FileName;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(txtCodigo.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string codigotexto = "";

            Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
            Workbook newWorkBook = appExcel.Workbooks.Open(txtDireccionArchivo.Text, true, true);
            _Worksheet objSheet = (_Worksheet)appExcel.ActiveWorkbook.ActiveSheet;
            object valores = new object();
            valores = "";
            //Colunas a exportar
            object Nombres = "B";
            object ApellidosPaterno = "C";
            object ApellidosMaterno = "D";
            object Email = "F";

            object Codigo = "A";

            object DNI = "E";


            //Inicio de a fina
            int InicioFila = 2;
            // Consulta MYSQL

          

            for (int i = InicioFila; i < 10500; i++)
            {
                //valores = objSheet.get_Range("C" + i).get_Value();
                Codigo = objSheet.get_Range("A" + i).get_Value();
                Nombres = objSheet.get_Range("B" + i).get_Value();
                ApellidosPaterno = objSheet.get_Range("C" + i).get_Value();
                ApellidosMaterno = objSheet.get_Range("D" + i).get_Value();
                DNI = objSheet.get_Range("E" + i).get_Value();
                Email = objSheet.get_Range("F" + i).get_Value();

                //DNI = DNI.ToString().Trim();

                String dnii = DNI + "";
                dnii = dnii.Trim();

                if (dnii.Length < 4)
                {
                    dnii = Codigo + "";
                }

                String codigoo = Codigo + "";
                codigoo = codigoo.Trim();


                string nomnrecompleto = Nombres + " " + ApellidosPaterno + " " + ApellidosMaterno;
                nomnrecompleto=nomnrecompleto.Replace('ñ','n');
                nomnrecompleto = nomnrecompleto.Replace('Ñ', 'N');
                nomnrecompleto = nomnrecompleto.Replace('á', 'a');
                nomnrecompleto = nomnrecompleto.Replace('é', 'e');
                nomnrecompleto = nomnrecompleto.Replace('í', 'i');
                nomnrecompleto = nomnrecompleto.Replace('ó', 'o');
                nomnrecompleto = nomnrecompleto.Replace('ú', 'u');
                nomnrecompleto = nomnrecompleto.Replace('Á', 'A');
                nomnrecompleto = nomnrecompleto.Replace('É', 'E');
                nomnrecompleto = nomnrecompleto.Replace('Í', 'I');
                nomnrecompleto = nomnrecompleto.Replace('Ó', 'O');
                nomnrecompleto = nomnrecompleto.Replace('Ú', 'U');
                nomnrecompleto = nomnrecompleto.Replace(' ', ' ');
                nomnrecompleto = nomnrecompleto.Replace('Ü', 'U');
                 



                string consulta = "INSERT INTO `users` (`id`, `username`, `password`, `name`, `email`, `company_name`, `company_tagline`, `company_website`, `company_logo`, `company_logo_dir`, `is_admin`, `hash_change_password`, `is_subscribed`, `expiry_date`, `created`, `modified`) VALUES (NULL, '" + codigoo + "', '" + dnii + "', '" + nomnrecompleto + "', '" + codigoo + "', 'UNAMBA', 'UNAMBA', '', '', '', '0', NULL, '1', NULL, NULL, NULL);";

                codigotexto = codigotexto + consulta + "\n";



                //System.Console.WriteLine("This is the value in the cell A1 - " + valores);
            }




            txtCodigo.Text = codigotexto;




        }
    }
}
