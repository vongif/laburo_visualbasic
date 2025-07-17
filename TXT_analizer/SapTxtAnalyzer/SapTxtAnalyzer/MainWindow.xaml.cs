using Microsoft.Win32;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace SapTxtAnalyzer
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private DataTable datosOriginales;

        private void CargarArchivos_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialogo = new OpenFileDialog
            {
                Filter = "Archivos de texto (*.txt)|*.txt",
                Multiselect = true
            };

            if (dialogo.ShowDialog() == true)
            {
                var tabla = new DataTable();

                foreach (string archivo in dialogo.FileNames)
                {
                    var lineas = File.ReadAllLines(archivo);
                    if (lineas.Length < 2) continue;

                    var encabezados = lineas[0].Split('\t');

                    // Solo crear columnas una vez
                    if (tabla.Columns.Count == 0)
                    {
                        foreach (var header in encabezados)
                            tabla.Columns.Add(header);
                        tabla.Columns.Add("FechaIngreso"); // NUEVA columna
                    }

                    foreach (var linea in lineas.Skip(1))
                    {
                        var campos = linea.Split('\t');
                        while (campos.Length < encabezados.Length)
                            campos = campos.Append("").ToArray();

                        var filaCompleta = campos.Take(encabezados.Length).Append(DateTime.Now.ToString("yyyy-MM-dd")).ToArray();
                        tabla.Rows.Add(filaCompleta);
                    }
                }

                datosOriginales = tabla;
                dgDatos.ItemsSource = datosOriginales.DefaultView;
                LlenarFiltros();
            }
        }

        private void LlenarFiltros()
        {
            cbCliente.ItemsSource = datosOriginales.AsEnumerable()
                .Select(r => r.Field<string>("Cliente")).Distinct().OrderBy(x => x).ToList();

            cbClasePedido.ItemsSource = datosOriginales.AsEnumerable()
                .Select(r => r.Field<string>("Clase de pedido")).Distinct().OrderBy(x => x).ToList();

            cbReferencia.ItemsSource = datosOriginales.AsEnumerable()
                .Select(r => r.Field<string>("Referencia cliente (OC)")).Distinct().OrderBy(x => x).ToList();

            cbFechaEntrega.ItemsSource = datosOriginales.AsEnumerable()
                .Select(r => r.Field<string>("Fecha Entrega")).Distinct().OrderBy(x => x).ToList();

            cbMaterial.ItemsSource = datosOriginales.AsEnumerable()
                .Select(r => r.Field<string>("Material")).Distinct().OrderBy(x => x).ToList();

            cbFechaIngreso.ItemsSource = datosOriginales.AsEnumerable()
                .Select(r => r.Field<string>("FechaIngreso")).Distinct().OrderBy(x => x).ToList();
        }



        private void Filtro_Changed(object sender, SelectionChangedEventArgs e)
        {
            var filtros = new List<string>();

            if (cbCliente.SelectedItem != null)
                filtros.Add($"[Cliente] = '{cbCliente.SelectedItem}'");

            if (cbClasePedido.SelectedItem != null)
                filtros.Add($"[Clase de pedido] = '{cbClasePedido.SelectedItem}'");

            if (cbReferencia.SelectedItem != null)
                filtros.Add($"[Referencia cliente (OC)] = '{cbReferencia.SelectedItem}'");

            if (cbFechaEntrega.SelectedItem != null)
                filtros.Add($"[Fecha Entrega] = '{cbFechaEntrega.SelectedItem}'");

            if (cbMaterial.SelectedItem != null)
                filtros.Add($"[Material] = '{cbMaterial.SelectedItem}'");

            if (cbFechaIngreso.SelectedItem != null)
                filtros.Add($"[FechaIngreso] = '{cbFechaIngreso.SelectedItem}'");

            datosOriginales.DefaultView.RowFilter = string.Join(" AND ", filtros);
        }





    }
}

