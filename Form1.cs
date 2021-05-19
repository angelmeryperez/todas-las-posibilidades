using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace todas_las_posibilidades
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        int g = 0, w1 = 0, w2 = 0, K = 0, w3 =0;
        bool vs = true;
        bool vc = true;
        string[] data  = new string[2];
        string[] data1 = new string[2];
        string validacion;

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //timer1.Start();
            dataGridView1.Rows.Add(9999);
            dataGridView2.Rows.Add(10000);
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                //DE ESTA MANERA FILTRAMOS TODOS LOS ARCHIVOS EXCEL EN EL NAVEGADOR DE ARCHIVOS
                Filter = "Excel | *.xls;*.xlsx;",

                //AQUÍ INDICAMOS QUE NOMBRE TENDRÁ EL NAVEGADOR DE ARCHIVOS COMO TITULO
                Title = "Seleccionar Archivo"
            };

            //EN CASO DE SELECCIONAR EL ARCHIVO, ENTONCES PROCEDEMOS A ABRIR EL ARCHIVO CORRESPONDIENTE
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                dataGridView3.DataSource = ImportarDatos(openFileDialog.FileName);
            }

            for (int i = 0, j = 0; i <= 99; i++)
            {
                for (int f = 0; f <= 99; f++, g++, j++)
                {
                    dataGridView1.Rows[j].Cells[0].Value = i;
                    dataGridView1.Rows[g].Cells[1].Value = f;
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

        }

        DataView ImportarDatos(string nombrearchivo) //COMO PARAMETROS OBTENEMOS EL NOMBRE DEL ARCHIVO A IMPORTAR
        {

            //UTILIZAMOS 12.0 DEPENDIENDO DE LA VERSION DEL EXCEL, EN CASO DE QUE LA VERSIÓN QUE TIENES ES INFERIOR AL DEL 2013, CAMBIAR A EXCEL 8.0 Y EN VEZ DE
            //ACE.OLEDB.12.0 UTILIZAR LO SIGUIENTE (Jet.Oledb.4.0)
            string conexion = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {0}; Extended Properties = 'Excel 12.0;'", nombrearchivo);
            nombrearchivo = "Regristro Actualizado";
            OleDbConnection conector = new OleDbConnection(conexion);

            conector.Open();

            //DEPENDIENDO DEL NOMBRE QUE TIENE LA PESTAÑA EN TU ARCHIVO EXCEL COLOCAR DENTRO DE LOS []
            OleDbCommand consulta = new OleDbCommand("select * from [Hoja1$]", conector);

            OleDbDataAdapter adaptador = new OleDbDataAdapter
            {
                SelectCommand = consulta
            };

            DataSet ds = new DataSet();

            adaptador.Fill(ds);

            conector.Close();

            return ds.Tables[0].DefaultView;


        }
        private void ExportarDatos(DataGridView datalistado)
        {

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application(); // Instancia a la libreria de Microsoft Office
            excel.Application.Workbooks.Add(true); //Con esto añadimos una hoja en el Excel para exportar los archivos
            int IndiceColumna = 0;
            foreach (DataGridViewColumn columna in datalistado.Columns) //Aquí empezamos a leer las columnas del listado a exportar
            {
                IndiceColumna++;
                excel.Cells[1, IndiceColumna] = columna.Name;
            }
            int IndiceFila = 0;
            foreach (DataGridViewRow fila in datalistado.Rows) //Aquí leemos las filas de las columnas leídas
            {
                IndiceFila++;
                IndiceColumna = 0;
                foreach (DataGridViewColumn columna in datalistado.Columns)
                {
                    IndiceColumna++;
                    excel.Cells[IndiceFila + 1, IndiceColumna] = fila.Cells[columna.Name].Value;
                }
            }
            excel.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            while (vs)
            {
                data[0] = dataGridView1.Rows[w1].Cells[0].Value.ToString();
                data[1] = dataGridView1.Rows[w1].Cells[1].Value.ToString();
                while (vc)
                {
                    
                    data1[0] = dataGridView3.Rows[w2].Cells[0].Value.ToString();
                    data1[1] = dataGridView3.Rows[w2].Cells[1].Value.ToString();

                    if ( (data[0] == data1[0] && data[1] == data1[1]) || (data[0] == data1[1] && data[1] == data1[0]))
                    {
                        K = 8;
                    }
                    w2++;
                    if (data1[0]== "1000") { vc= false; }
                }
                w2 = 0;
                if (K < 1)
                {
                    dataGridView2.Rows[w3].Cells[0].Value = data[0];
                    dataGridView2.Rows[w3].Cells[1].Value = data[1];
                    w3++;
                    K = 0;
                }
                
                if (data[0] == "99" && data[1] == "99") { vs = false; }
                w1++;
            }
        }
    }
}
