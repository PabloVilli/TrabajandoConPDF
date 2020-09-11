using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Text.RegularExpressions;

namespace LectorPDF_V01_dsa
{
    public partial class Form1 : Form
    {
        //Variables para guardar el UUID Y el MX.
        string _uuid, _mx;
        //Variable que guardara la ruta.
        string ruta;
        //Creacion de la tabla.
        DataTable table = new DataTable("DatosFactura");
        //creacion de una columna.
        DataColumn column;
        //creacion de una fila.
        DataRow row;
        //Esta expresion analiza un UUID conformado por ceros. 
        string exp = "^[0]{8}-[0]{4}-[0]{4}-[0]{4}-[0]{12}$";
        //Esta expresion analiza un UUID valido. 
        string exp_uuid = "^[A-Z|0-9]{8}-[A-Z|0-9]{4}-[A-Z|0-9]{4}-[A-Z|0-9]{4}-[A-Z|0-9]{12}$";
        //Esta expresion analiza un MX valido.
        string exp_mx = "^MX ?[0-9]{6}$";
        //Esta variable guarda el resultado de la expresion si es que se encuentra una coincidencia.
        string a;

        public Form1()
        {
            InitializeComponent();
            //Inicializo los dos string.
            _uuid = "uuid";
            _mx = "-mx-";
        }

        //Creo la tabla donde guardare los datos. 
        public void createDataTable()
        {
            //Creo la columna para UUID
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "UUID";
            table.Columns.Add(column);

            //Creo la columna para MX
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "MX";
            table.Columns.Add(column);
        }

        //Este metodo lee el PDF
        public string ReadPdfFile(object Filename)
        {
            string strText = string.Empty;
            //try para obtener el error en caso de que ocurra
            try
            {
                PdfReader readerPdf = new PdfReader((string)Filename);

                for (int page = 1; page <= readerPdf.NumberOfPages; page++)
                {
                    ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy();
                    PdfReader reader = new PdfReader((string)Filename);
                    String s = PdfTextExtractor.GetTextFromPage(reader, page, its);

                    s = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(s)));
                    strText = strText + s;
                    reader.Close();
                }
            }
            //Manejo de errores
            catch (Exception ex)
            {
                //Muestro el error en caso de que ocurra
                lblMsjs.ForeColor = Color.Crimson;
                lblMsjs.Text = ex.Message.ToString();
            }
            //retorno el texto 
            return strText;
        }

        //Este metodo recibe un texto y tambien la expresion que buscara dentro del texto.
        public string MatchesExp(string texto, string exp)
        {
            MatchCollection find = Regex.Matches(texto, exp);
            //buscara por cada linea la expresion buscada.
            foreach (Match e in find)
            {
                a = e.ToString();
            }
            //Devuelve el texto que coicidio 
            return a;
        }

        //Boton para leer el PDF.
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            //Este label muestra mensajes en la pantalla, aqui limpio el texto por si tiene.
            lblMsjs.Text = string.Empty;
            //Limpio la tabla para que no muestre datos de ls anteriores
            table.Clear();
            //Limpio del origen de datos del grid
            dgv2.DataSource = null;
            //Esta variable recibe el texto que se obtuvo del PDF. 
            string salida;

            //Abro una ventana nueva del explorador de y para seleccionar ahivos.
            OpenFileDialog openArc = new OpenFileDialog();
            //Le indico que buscara solo archivos de tio 
            openArc.Filter = "PDF Files | *.pdf";
            //Obtiene o establece un valor que indica si el cuadro de diálogo muestra una advertencia cuando el usuario 
            //especifica un nombre de archivo que no existe.
            openArc.CheckFileExists = true;
            //Obtiene o establece un valor que indica si el cuadro de diálogo permite seleccionar varios archivos.
            openArc.Multiselect = true;
            //Si el usuario da clic en aceptar 
            if (openArc.ShowDialog() == DialogResult.OK)
            {
                //Se muestra un panel con el grid
                panel1.Visible = true;
                //Recorro todos los
                for (int i = 0; i < openArc.SafeFileNames.Count(); i++)
                {
                    //Creo una variable que guarde los caracteres que busco dentro del nombre y ruta de los arhivos.
                    string b = "\\";
                    //Busco dentro del nombre del archivo la poosicion que me indique la ultima aparicion de la variable b y la guardo en una variable int.
                    int position = openArc.FileName.LastIndexOf(b);
                    //Aqui inicializo la variable con la ruta del archivo pero elimino el nombre del archivo.
                    ruta = @"" + openArc.FileName.Substring(0, position);
                    //Guardo el texto que me devuelve el metodo de Leer un PDF.
                    salida = ReadPdfFile(ruta + b + openArc.SafeFileNames[i]);
                    //Guardo el texto que esta en la variable salida pero al mismo tiempo lo separa por saltos de linea y cada linea es un elemento distinto dentro del arregle.
                    string[] arregloString = salida.Split('\n');
                    //Inicializo una nueva fila para la tabla
                    row = table.NewRow();
                    //Cada elemento del arreglo sera evaluado con la expresion del UUID y si la expresión existe sera almacenada en una variable.
                    foreach (string item in arregloString)
                    {
                        _uuid = MatchesExp(item, exp_uuid);
                    }
                    //Cada elemento del arreglo sera evaluado con la expresion del MX y si la expresión existe sera almacenada en una variable.
                    foreach (string item in arregloString)
                    {
                        _mx = MatchesExp(item, exp_mx);
                    }
                    //Agrego las variables si es que se enctontro alguna coincidencia.
                    row["UUID"] = _uuid;
                    row["MX"] = _mx;
                    //Agrego valores a las columnas de las fila.
                    table.Rows.Add(row);
                    //Le mando la tabla con datos al gridView.
                    dgv2.DataSource = table;
                    //Agrego un mesaje al fnalizar.
                    lblMsjs.ForeColor = Color.Blue;
                    lblMsjs.Text = "Success";

                }
            }
        }

        //Boton para salir de la aplicación.
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        //Aqui inicializo la tabla.
        private void Form1_Load(object sender, EventArgs e)
        {
            createDataTable();

        }
    }
}
