using Indexeo_Proyectos_Optimizado.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Indexeo_Proyectos_Optimizado
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        String AltaDocClientes="", Atencioncliente="", CierreProyecto="", ComprasProyecto="", ComprasFacturas="", ComprasOrdendeCompra="",ComprasComparativas="", DosierCertificadoCalidad="", DosierEquipos="",ListaInox="",InstalacionInox="",TotalComun="";
        String DosierCartaGarantia, DosierConstancia, DosierFotos, DoiserCalidadAgua, DosierPruebasEquipo, DosierPruebasEstanquiedad="", DosierPruebasInstalaciones="", Encuestas="", EntregaDocumentos="", ExpedienteComercial="", ExpedienteTecnico="", FabricacionyEquipamiento="", InstalacionesyArranque="", InstalacionAlmacen="", InstalacionCajachica="", InstalacionComprobaciones="", InstalacionDispersion="", InstalacionFotos="", InstalacionesOrden="", InstalacionReportes="", InstalacioneRequisiciones="", OperacionyMantenimiento="", PedidoInterno="", Planos="", Posventa="",Obracivil,ListadeMateriales="",InstalacionyPreparativos="";

        decimal porcentaje2;
        String ValorLista = "0", ValorPreparativos = "0", NombreProyecto="";
        String ObraCivil = "", SirocEIMSS = "", Subcontratos = "", ControldeObra = "", TotalComunOtro = "", Presupuesto = "";
        decimal contador1, contador2, contador3, contador4, contador5, contador6, contador7, contador8, contador9, contador10,
           contador11, contador12, contador13, contador14, contador15, contador16, contador17, contador18, contador19, contador20, contador21, contador22, contador23, contador24,contador25,contador26,contador27,contador28,contador29,contador30;
        decimal porcentaje, generalporcentaje, porcentajeinox;
        decimal valorproyecto;
        String TotalProyectos = "",Año,Tipo,Nombre,Folio,Totalrestantes,Documento,Carpeta,Departamento;
        int TotalProyectosEntero, contadorProyectos, k1, Totalrestantesentero,k2,contadorRestantes;
        string[] files;
        int i;
        string
AtencionClientes
, AtencionClientes_Atencion
, AtencionClientes_EntregaDocumentos
, AtencionClientes_Postventa
, AtencionClientes_Encuesta
, Finanzas
, Finanzas_CierreProyecto
, Compraas
, DosierdeCalidad
, Ventas
, Ventas_AltayDocumentosClientes
, Ventas_PedidoInterno
, Ventas_ExpedienteComercial
, Proyectos
, Proyectos_ExpedienteTecnic
, Proyectos_Planos
, EquipamientoeInstalaciones
, EquipamientoeInstalaciones_FabricacionyEquipamiento
, EquipamientoeInstalaciones_Instalacionyarranque
, EquipamientoeInstalaciones_OperacionyMantenimiento
, General;

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: esta línea de código carga datos en la tabla 'proyectos_Avance_Others.Proyectos_Avance_other' Puede moverla o quitarla según sea necesario.
            this.proyectos_Avance_otherTableAdapter.Fill(this.proyectos_Avance_Others.Proyectos_Avance_other);
            // TODO: esta línea de código carga datos en la tabla 'proyectos_Avance_INOX._Proyectos_Avance_INOX' Puede moverla o quitarla según sea necesario.
            this.proyectos_Avance_INOXTableAdapter.Fill(this.proyectos_Avance_INOX._Proyectos_Avance_INOX);
            // TODO: esta línea de código carga datos en la tabla 'proyectosArchivosInox.Ser_Documentos_INOX' Puede moverla o quitarla según sea necesario.
            this.ser_Documentos_INOXTableAdapter.Fill(this.proyectosArchivosInox.Ser_Documentos_INOX);
            // TODO: esta línea de código carga datos en la tabla 'proyectosArchivos.Proyectos_Avance_Archivos' Puede moverla o quitarla según sea necesario.
            this.proyectos_Avance_ArchivosTableAdapter.Fill(this.proyectosArchivos.Proyectos_Avance_Archivos);
            // TODO: esta línea de código carga datos en la tabla 'proyectos_Avance._Proyectos_Avance' Puede moverla o quitarla según sea necesario.
            this.proyectos_AvanceTableAdapter.Fill(this.proyectos_Avance._Proyectos_Avance);
            proyectos_AvanceTableAdapter.LimpiadorPorcentajes();
            iniciador();
        }

        public static string ObtenerCadena()
        {
            return Settings.Default.CBR_IngenieriaConnectionString;/*Este codigo obtiene los datos de la cadena de conexion declarada en los settings de la aplicacion*/

        }
        public void iniciador()
        {

            SqlConnection conexion = new SqlConnection(ObtenerCadena());
            conexion.Open();
            ////////////////////////////////////////////////////////////////////////////////////
            ////////////////////////////////////SE EJECUTA EL COMANDO HACIA LA VISTA Y SE HACE LA VALIDACION//////////////////////////////////////
            SqlCommand cmd = new SqlCommand(
                                            "select " +
                                            "Count('Consecutivo') " +
                                            "from [Folio_proyectos>2019] "

                                            , conexion);


            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.SelectCommand.CommandTimeout = 126000;
            DataTable dt = new DataTable();
            sda.Fill(dt);
            conexion.Dispose();
            if (dt.Rows.Count > 0)
            {
                TotalProyectos = dt.Rows[0][0].ToString();
                Contador();
                conexion.Dispose();
            }
            else { }
            conexion.Dispose();
        }

        public void Contador()
        {

            SqlConnection conexion = new SqlConnection(ObtenerCadena());

            TotalProyectosEntero = Int32.Parse(TotalProyectos);

            contadorProyectos = 0;
            conexion.Dispose();
            for (k1 = 1; k1 <= TotalProyectosEntero; k1++)
            {
                contadorProyectos = contadorProyectos + 1;

                ConsultaProyecto();


                ContadorRestantes();
                ListadeMateriales = "";
                contador1 = 0; contador2 = 0; contador3 = 0; contador4 = 0; contador5 = 0; contador6 = 0; contador7 = 0; contador8 = 0; contador9 = 0; contador10 = 0;
                contador11 = 0; contador12 = 0; contador13 = 0; contador14 = 0; contador15 = 0; contador16 = 0; contador17 = 0; contador18 = 0; contador19 = 0; contador20 = 0; contador21 = 0; contador22 = 0;

            }

            conexion.Dispose();
        }
        public void ConsultaProyecto()
        {

            if (Tipo == "Proyectos-Inox")
            {

                ActualizaPorcentajesINOXFinal();
            }
            else if (Tipo == "" || Tipo == null) { }
            else
            {

                ActualizaPorcentajesINDPotFinal();
            }
           

            SqlConnection conexion = new SqlConnection(ObtenerCadena());
            conexion.Open();
            SqlCommand cmd = new SqlCommand(
                                   "select top " + "(" + contadorProyectos + ")" +
                                   "[Ano], " +
                                   "[tipo2], " +
                                   "[Nombre], " +
                                   "[Nombre2] " +

                                   "from [Folio_proyectos>2019] "

                                   , conexion);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.SelectCommand.CommandTimeout = 36000;
            DataTable dt = new DataTable();
            sda.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                Año = dt.Rows[k1 - 1][0].ToString();
                Tipo = dt.Rows[k1 - 1][1].ToString();
                Folio = dt.Rows[k1 - 1][2].ToString();
                Nombre = dt.Rows[k1 - 1][3].ToString();
                NombreProyecto = dt.Rows[k1 - 1][3].ToString();
                Thread.Sleep(50);

                conexion.Dispose();

                files = Directory.GetFiles(@"G:\SGC-PROYECTOS-CBR\SGC\" + Año + "\\" + Tipo + "\\" + Folio, "*", SearchOption.AllDirectories);
                index_carpetasTableAdapter.limpia();
                Thread.Sleep(50);
                porcentaje2 = 0;

                if (Tipo == "Proyectos-Inox")
                {
                
                    index_carpetasTableAdapter.InsertQueryINOX();
                }
                else {
                   
                    index_carpetasTableAdapter.InsertQueryPOT(); }
                index_carpetasTableAdapter.InsertQuery();
    
                Ingresor();
                i = 0;
               
                for ( i = 0; i < files.Length; i++)
                {
                    comboBox1.Items.Add(Path.GetFileName(files[i]));

                    SqlConnection conexion2 = new SqlConnection(ObtenerCadena());
                    conexion2.Open();
                    SqlCommand cmd2 = new SqlCommand(
                                           "select " +
        "  [Nombre]" +



                                           "from [Index_carpetas] where Nombre = @Nombre"

                                           , conexion2);
                    SqlDataAdapter sda2 = new SqlDataAdapter(cmd2);
                    cmd2.Parameters.AddWithValue("Nombre", Path.GetFileName(files[i]));
                    sda2.SelectCommand.CommandTimeout = 136000;
                    DataTable dt2 = new DataTable();
                    sda2.Fill(dt2);
                    conexion2.Dispose();
                    if (dt2.Rows.Count > 0)
                    {
                        ActualizaPorcentajes();

                      
                        if (Tipo == "Proyectos-Inox")
                        {
                            ActualizaPorcentajesINOX();
                        }
                        else { ActualizaPorcentajesINDPot(); }
                    }
                   

                        index_carpetasTableAdapter.DeleteQuery(Path.GetFileName(files[i]));

                }

            }
            else { }
           
        }
        public void Ingresor() {

            SqlConnection conexion2 = new SqlConnection(ObtenerCadena());
            conexion2.Open();
            SqlCommand cmd2 = new SqlCommand(
                                   "select " +
"  [AtencionClientes]" +
" ,[AtencionClientes_Atencion]" +
" ,[AtencionClientes_EntregaDocumentos]" +
" ,[AtencionClientes_Postventa]" +
" ,[AtencionClientes_Encuesta]" +
" ,[Finanzas]" +
" ,[Finanzas_CierreProyecto]" +
" ,[Compras]" +
" ,[Dosier de Calidad]" +
" ,[Ventas]" +
" ,[Ventas_AltayDocumentosClientes]" +
" ,[Ventas_PedidoInterno]" +
" ,[Ventas_ExpedienteComercial]" +
" ,[Proyectos]" +
" ,[Proyectos_ExpedienteTecnico]" +
" ,[Proyectos_Planos]" +
" ,[EquipamientoeInstalaciones]" +
" ,[EquipamientoeInstalaciones_Fabricacion y Equipamiento]" +
" ,[EquipamientoeInstalaciones_Instalacionyarranque]" +
" ,[EquipamientoeInstalaciones_OperacionyMantenimiento]" +
" ,[General]" +


                                   "from [Proyectos_Avance] where Proyecto = @proyecto"

                                   , conexion2);
            SqlDataAdapter sda2 = new SqlDataAdapter(cmd2);
            cmd2.Parameters.AddWithValue("proyecto", Nombre);
            sda2.SelectCommand.CommandTimeout = 36000;
            DataTable dt2 = new DataTable();
            sda2.Fill(dt2);
            if (dt2.Rows.Count > 0)
            {
                AtencionClientes = dt2.Rows[0][0].ToString();
                AtencionClientes_Atencion = dt2.Rows[0][1].ToString();
                AtencionClientes_EntregaDocumentos = dt2.Rows[0][2].ToString();
                AtencionClientes_Postventa = dt2.Rows[0][3].ToString();
                AtencionClientes_Encuesta = dt2.Rows[0][4].ToString();
                Finanzas = dt2.Rows[0][5].ToString();
                Finanzas_CierreProyecto = dt2.Rows[0][6].ToString();
                Compraas = dt2.Rows[0][7].ToString();
                DosierdeCalidad = dt2.Rows[0][8].ToString();
                Ventas = dt2.Rows[0][9].ToString();
                Ventas_AltayDocumentosClientes = dt2.Rows[0][10].ToString();
                Ventas_PedidoInterno = dt2.Rows[0][11].ToString();
                Ventas_ExpedienteComercial = dt2.Rows[0][12].ToString();
                Proyectos = dt2.Rows[0][13].ToString();
                Proyectos_ExpedienteTecnic = dt2.Rows[0][14].ToString();
                Proyectos_Planos = dt2.Rows[0][15].ToString();
                EquipamientoeInstalaciones = dt2.Rows[0][16].ToString();
                EquipamientoeInstalaciones_FabricacionyEquipamiento = dt2.Rows[0][17].ToString();
                EquipamientoeInstalaciones_Instalacionyarranque = dt2.Rows[0][18].ToString();
                EquipamientoeInstalaciones_OperacionyMantenimiento = dt2.Rows[0][19].ToString();
                General = dt2.Rows[0][20].ToString();
                conexion2.Dispose();
                //proyectos_Avance_INOXTableAdapter.UpdateTotalComun(General,Nombre);
            }
         
            else
            {

                proyectos_AvanceTableAdapter.AgregaProyecto(Nombre, "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0");
                if (Tipo == "Proyectos-Inox")
                {
                    proyectos_Avance_INOXTableAdapter.InsertQuery(Nombre, "0", "0", "0", "0");
                }
                else
                {
                    proyectos_Avance_otherTableAdapter.InsertQuery(Nombre, "0", "0", "0", "0", "0","0","0");
                }
                conexion2.Dispose();

            }
     
        }

        public void ContadorRestantes()
        {

            SqlConnection conexion = new SqlConnection(ObtenerCadena());
            conexion.Open();
            ////////////////////////////////////////////////////////////////////////////////////
            ////////////////////////////////////SE EJECUTA EL COMANDO HACIA LA VISTA Y SE HACE LA VALIDACION//////////////////////////////////////
            SqlCommand cmd = new SqlCommand(
                                            "select " +
                                            "Count('Consecutivo') " +
                                            "from [Index_carpetas] "

                                            , conexion);


            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.SelectCommand.CommandTimeout = 36000;
            DataTable dt = new DataTable();
            sda.Fill(dt);
            conexion.Dispose();
            if (dt.Rows.Count > 0)
            {
                Totalrestantes = dt.Rows[0][0].ToString();
                Faltantes();
            }
            else { }

        }
        public void Faltantes()
        {
            Totalrestantesentero = 0;
            AltaDocClientes = "";
            Atencioncliente = "";
            CierreProyecto = "";
            ComprasProyecto = ""; ComprasFacturas = ""; ComprasOrdendeCompra = ""; ComprasComparativas = ""; DosierCertificadoCalidad = ""; DosierEquipos = "";
            DosierCartaGarantia = ""; DosierConstancia = ""; DosierFotos=""; DoiserCalidadAgua = ""; DosierPruebasEquipo = ""; DosierPruebasEstanquiedad = ""; DosierPruebasInstalaciones = ""; Encuestas = ""; EntregaDocumentos = ""; ExpedienteComercial = ""; ExpedienteTecnico = ""; FabricacionyEquipamiento = ""; InstalacionesyArranque = ""; InstalacionAlmacen = ""; InstalacionCajachica = ""; InstalacionComprobaciones = ""; InstalacionDispersion = ""; InstalacionFotos = ""; InstalacionesOrden = ""; InstalacionReportes = ""; InstalacioneRequisiciones = ""; OperacionyMantenimiento = ""; PedidoInterno = ""; Planos = ""; Posventa = "";
            ValorLista = "";
            ValorPreparativos = "";
               Totalrestantesentero = Int32.Parse(Totalrestantes);

            contadorRestantes = 0;

            for (k2 = 1; k2 <= Totalrestantesentero; k2++)
            {
             
                if (contadorRestantes <= Totalrestantesentero)
                {
                  
                     contadorRestantes = contadorRestantes + 1;
                    ListaFaltantes();
                }
                else { }
            }
          

        }
        public void ListaFaltantes()
        {
            SqlConnection conexion2 = new SqlConnection(ObtenerCadena());
            conexion2.Open();
            SqlCommand cmd2 = new SqlCommand(
                                   "select " +



"  [Proyecto]" +



                                   "from [Proyectos_Avance_Archivos] where Proyecto = @Nombre"

                                   , conexion2);
            SqlDataAdapter sda2 = new SqlDataAdapter(cmd2);
            cmd2.Parameters.AddWithValue("Nombre", Nombre);
            sda2.SelectCommand.CommandTimeout = 136000;
            DataTable dt2 = new DataTable();
            sda2.Fill(dt2);
            conexion2.Dispose();
            if (dt2.Rows.Count > 0)
            {  }

         

        

    
            else { proyectos_Avance_ArchivosTableAdapter.InsertQuery(Nombre, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""); }





SqlConnection conexion = new SqlConnection(ObtenerCadena());
            conexion.Open();





            SqlCommand cmd = new SqlCommand(
                                   "select top " + "(" + contadorRestantes + ")" +
                                     "[Departamento], " +
                                   "[Carpeta], " +
                                   "[Nombre] " +

                                   "from [Index_carpetas] "

                                   , conexion);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);

            sda.SelectCommand.CommandTimeout = 36000;
            DataTable dt = new DataTable();
            sda.Fill(dt);
            conexion.Dispose();
            if (dt.Rows.Count > 0)
            {
             
               
                Documento = dt.Rows[contadorRestantes-1][2].ToString();
                Carpeta = dt.Rows[contadorRestantes-1][1].ToString();
                Departamento = dt.Rows[contadorRestantes-1][0].ToString();
                Thread.Sleep(50);
                String validadorinox = "",validadorother="";

                if (Directory.Exists("G:SGC-PROYECTOS-CBR/SGC/" + Año + "/" + Tipo + "/"+Folio+"/Lista de Materiales y Equipos"))
                {
                    validadorinox = "ok";
                }
                else
                {
                    validadorinox = "na";
                    proyectos_Avance_ArchivosTableAdapter.Update_ListadeMateriales("NA", Nombre);
                    proyectos_Avance_ArchivosTableAdapter.Update_Instalacionypreparativos("NA", Nombre);
                }
                if (Directory.Exists("G:SGC-PROYECTOS-CBR/SGC/" + Año + "/" + Tipo + "/" + Folio + "/Obra civil"))
                {
                    validadorother = "ok";
                }
                else
                {
                    validadorother = "na";
                    proyectos_Avance_ArchivosTableAdapter.Update_Obracivil("NA", Nombre);
                }
                switch (Carpeta)
                {

                    case "Alta y Documentos Clientes":
                        AltaDocClientes = AltaDocClientes + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_AltaDocumentosCliente(AltaDocClientes, Nombre);
                        break;
                    case "Encuestas de Satisfaccion":
                        Encuestas = Encuestas + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Encuesta(Encuestas, Nombre);
                        break;

                    case "Atencion Cliente":
                        Atencioncliente = Atencioncliente + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Atencionaclientes(Atencioncliente, Nombre);
                        break;
                    case "Cierre Proyecto":
                        CierreProyecto = CierreProyecto + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_FinanzasCierre(CierreProyecto, Nombre);
                        break;
                    case "Compras/Comparativas":
                        ComprasComparativas = ComprasComparativas + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Compras(ComprasComparativas, Nombre);
                        break;
                    case "Compras/Facturas":
                        ComprasFacturas = ComprasFacturas + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Compras(ComprasFacturas, Nombre);
                        break;
                    case "Compras/Ordenes de Compra":
                        ComprasOrdendeCompra = ComprasOrdendeCompra + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Compras(ComprasOrdendeCompra, Nombre);
                        break;

                    case "DOSSIER DE CALIDAD/Certificados de calidad":
                        DosierdeCalidad = DosierdeCalidad + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;

                    case "DOSSIER DE CALIDAD/Certificados de calidad/Equipos":
                        DosierdeCalidad = DosierdeCalidad + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;
                    case "DOSSIER DE CALIDAD/Cliente/Carta Garantia":
                        DosierdeCalidad = DosierdeCalidad + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;
                    case "DOSSIER DE CALIDAD/Cliente/Constancia":
                        DosierdeCalidad = DosierdeCalidad + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;
                    case "DOSSIER DE CALIDAD/Fotografias/Fotos":
                        DosierdeCalidad = DosierdeCalidad + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;
                    case "DOSSIER DE CALIDAD/Pruebas de Calidad/Analisis Agua":
                        DosierdeCalidad = DosierdeCalidad + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;

                    case "DOSSIER DE CALIDAD/Pruebas de Calidad/Equipos":
                        DosierdeCalidad = DosierdeCalidad + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;

                    case "DOSSIER DE CALIDAD/Pruebas de Calidad/Estanquiedad":
                        DosierdeCalidad = DosierdeCalidad + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;
                    case "DOSSIER DE CALIDAD/Pruebas de Calidad/Instalaciones":
                        DosierdeCalidad = DosierdeCalidad + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;
                    case "Encuestas satisfaccion":
                        Encuestas = Encuestas + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Encuesta(Encuestas, Nombre);
                        break;
                    case "Entrega Documentos":
                        EntregaDocumentos = EntregaDocumentos + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_EntregaDocumentos(EntregaDocumentos, Nombre);
                        break;
                    case "Expediente Comercial":
                        ExpedienteComercial = ExpedienteComercial + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_ExpedienteComercial(ExpedienteComercial, Nombre);
                        break;
                    case "Expediente Tecnico":
                        ExpedienteTecnico = ExpedienteTecnico + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_ExpedienteTecnico(ExpedienteTecnico, Nombre);
                        break;

                    case "Fabricacion y Equipamiento":
                        FabricacionyEquipamiento = FabricacionyEquipamiento + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Equipamiento_FabricacionyEquipamiento(FabricacionyEquipamiento, Nombre);
                        break;

                    case "Instalacion y arranque":
                        InstalacionesyArranque = InstalacionesyArranque + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Equipamiento_InstalacionyArranque(InstalacionesyArranque, Nombre);
                        break;
                    case "Instalacion y arranque/Almacen":
                        InstalacionAlmacen = InstalacionesyArranque + InstalacionAlmacen + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Equipamiento_InstalacionyArranque(InstalacionAlmacen, Nombre);
                        break;
                    case "Instalacion y arranque/Comprobaciones/Caja chica":
                        InstalacionCajachica = InstalacionesyArranque + InstalacionCajachica + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Equipamiento_InstalacionyArranque(InstalacionCajachica, Nombre);
                        break;
                    case "Instalacion y arranque/Comprobaciones/Comprobaciones":
                        InstalacionComprobaciones = InstalacionesyArranque + InstalacionComprobaciones + " " + Documento; ;
                        proyectos_Avance_ArchivosTableAdapter.Update_Equipamiento_InstalacionyArranque(InstalacionComprobaciones, Nombre);
                        break;
                    case "Instalacion y arranque/Dispersion":
                        InstalacionDispersion = InstalacionesyArranque + InstalacionDispersion + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Equipamiento_InstalacionyArranque(InstalacionDispersion, Nombre);
                        break;

                    case "Instalacion y arranque/Fotos":
                        InstalacionFotos = InstalacionesyArranque + InstalacionFotos + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Equipamiento_InstalacionyArranque(InstalacionFotos, Nombre);
                        break;

                    case "Instalacion y arranque/Ordenes de Cambio":
                        InstalacionesOrden = InstalacionesyArranque + InstalacionesOrden + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Equipamiento_InstalacionyArranque(InstalacionesOrden, Nombre);
                        break;
                    case "Instalacion y arranque/Reportes":
                        InstalacionReportes = InstalacionesyArranque + InstalacionReportes + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Equipamiento_InstalacionyArranque(InstalacionReportes, Nombre);
                        break;
                    case "Instalacion y arranque/Requisiciones":
                        InstalacioneRequisiciones = InstalacionesyArranque + InstalacioneRequisiciones + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Equipamiento_InstalacionyArranque(InstalacioneRequisiciones, Nombre);
                        break;
                    case "Operacion y Mantenimiento":
                        OperacionyMantenimiento = OperacionyMantenimiento + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_OperacionyMantenimiento(OperacionyMantenimiento, Nombre);
                        break;
                    case "Pedido Interno":
                        PedidoInterno = PedidoInterno + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_PedidoInterno(PedidoInterno, Nombre);
                        break;
                    case "Planos":
                        Planos = Planos + " " + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Planos(Planos + " ", Nombre);
                        break;
                    case "Postventa":
                        Posventa = Posventa + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_Postventa(Posventa + " ", Nombre);
                        break;
                        ///////////////////////////INOX////////////////////////////////////////////////////////
                        ///
                   

                    case "Lista de Materiales y Equipos":
                        if (validadorinox == "ok")
                        {
                            ListadeMateriales = ListadeMateriales + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_ListadeMateriales(ListadeMateriales + "\n", Nombre);
                        }
                        else {  }

                        break;


                    case "Instalacion y Preparativos":
                        if (validadorinox == "ok")
                        {
                            InstalacionyPreparativos = InstalacionyPreparativos + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Instalacionypreparativos(InstalacionyPreparativos + "\n", Nombre);
                        }
                        else { }
                        break;


                      
                        //////////////////////////INOX//////////////////////////////////////////////////////////
                        ///
                        ////////////////////OTher////////////////////////////////////////
                   
                    case "Obra civil/Presupuesto/Ctg Obra":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;
                    case "Obra civil/Subcontrato/Expediente de Obra":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;
                    case "Obra civil/Subcontrato/Programa de Obra":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else { }
                        break;
                    case "Obra civil/Subcontrato/Consolidados":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;

                    case "Obra civil/Control de Obra/Liberaciones":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;


                    case "Obra civil/Control de Obra/Requisiciones":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;


                    case "Obra civil/Control de Obra/Bitacora":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;

                    case "Obra civil/Control de Obra/Dispersiones":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else { ; }
                        break;

                    case "Obra civil/Control de Obra/Almacen":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else { }
                        break;

                    case "Obra civil/Control de Obra/Comprobacion Gastos":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;

                    case "Obra civil/Control de Obra/Reporte/Fotografico":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;

                    case "Obra civil/Control de Obra/Reporte/Fisico Fin":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;

                    case "Obra civil/Control de Obra/Reportes":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;
                    case "Obra civil/Control de Obra/Estimaciones":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;
                    case "Obra civil/Control de Obra/Reportes/Direccion":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else { }
                        break;
                    case "Obra civil/Control de Obra/Dosier de Calidad":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;
                    case "Obra civil/Control de Obra/Seguridad e Higiene":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;
                    case "Obra civil/Control de Obra/Orden de Cambios":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;
                    case "Obra civil/Control de Obra/fotos":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else { }
                        break;
                    case "Obra civil/Control de Obra/Lista de Materiales y equipos":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else { }
                        break;
                    case "Obra civil/SIROC E IMSS/Registro de Siroc":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;
                    case "Obra civil/SIROC E IMSS/Subcontratistas":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else { }
                        break;
                    case "Obra civil/SIROC E IMSS/Altas":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else { }
                        break;
                    case "Obra civil/SIROC E IMSS/Pagos Imss":
                      
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;
                    case "Obra civil/Subcontratos/Contrato":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else { ; }
                        break;
                    case "Obra civil/Subcontratos/Presupuesto":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else { }
                        break;
                    case "Obra civil/Subcontratos/Concentrado Estimaciones":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else {  }
                        break;
                    case "Obra civil/Subcontratos/Estimaciones":
                        if (validadorother == "ok")
                        {
                            Obracivil = Obracivil + "\n" + Documento;
                            proyectos_Avance_ArchivosTableAdapter.Update_Obracivil(Obracivil + "\n", Nombre);
                        }
                        else { }
                        break;
                    case "DOSSIER DE CALIDAD/Certificados Calidad Acero a carbon-fogo":
                        DosierdeCalidad = DosierdeCalidad + "\n" + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;
                    case "DOSSIER DE CALIDAD/Certificados Calidad Acero":
                        DosierdeCalidad = DosierdeCalidad + "\n" + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;
                    case "DOSSIER DE CALIDAD/Pruebas de Calidad Concreto":
                        DosierdeCalidad = DosierdeCalidad + "\n" + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;
                    case "DOSSIER DE CALIDAD/Pruebas de Calidad Compactacion":
                        DosierdeCalidad = DosierdeCalidad + "\n" + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;
                    case "DOSSIER DE CALIDAD/Revisiones":
                        DosierdeCalidad = DosierdeCalidad + "\n" + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;
                        break;
                    case "DOSSIER DE CALIDAD/Pruebas electricas":
                        DosierdeCalidad = DosierdeCalidad + "\n" + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;
                       
                    case "DOSSIER DE CALIDAD/Concreto laboratorio":
                        DosierdeCalidad = DosierdeCalidad + "\n" + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;
                    case "DOSSIER DE CALIDAD":
                        DosierdeCalidad = DosierdeCalidad + "\n" + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;
                  
                    case "DOSSIER DE CALIDAD/Certificados de Calidad":
                        DosierdeCalidad = DosierdeCalidad + "\n" + Documento;
                        proyectos_Avance_ArchivosTableAdapter.Update_DosierCalidad(DosierdeCalidad, Nombre);
                        break;



                    default:
                        break;

                        Documento = "";

                        //if (ObraCivil == "0" && SirocEIMSS == "0" && Subcontratos == "0" && ControldeObra == "0" && Presupuesto == "0")
                        //{
                        //    proyectos_Avance_otherTableAdapter.UpdateTotalComun(General, NombreProyecto);
                        //}
                        //else { }

                        //if (ValorLista == "0" && ValorPreparativos == "0")
                        //{
                        //    proyectos_Avance_INOXTableAdapter.UpdateTotalComun(General, NombreProyecto);
                        //}
                        //else { }
                }
            }
        }

        public void ActualizaPorcentajes()
        {

            SqlConnection conexion = new SqlConnection(ObtenerCadena());
            conexion.Open();





            SqlCommand cmd = new SqlCommand(
                                   "select" +
                                     "[Departamento], " +
                                   "[Carpeta], " +
                                   "[Documento] " +

                                   "from [Ser_Documentos] where Documento = @Nombre"

                                   , conexion);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            cmd.Parameters.AddWithValue("Nombre", Path.GetFileName(files[i]));
            sda.SelectCommand.CommandTimeout = 36000;
            DataTable dt = new DataTable();
            sda.Fill(dt);
            conexion.Dispose();
            if (dt.Rows.Count > 0)
            {
                Documento = dt.Rows[0][2].ToString();
                Carpeta = dt.Rows[0][1].ToString();
                Departamento = dt.Rows[0][0].ToString();
                Thread.Sleep(50);


            }
            SqlConnection conexion2 = new SqlConnection(ObtenerCadena());
            conexion2.Open();
            SqlCommand cmd2 = new SqlCommand(
                                   "select " +
"  [AtencionClientes]" +
" ,[AtencionClientes_Atencion]" +
" ,[AtencionClientes_EntregaDocumentos]" +
" ,[AtencionClientes_Postventa]" +
" ,[AtencionClientes_Encuesta]" +
" ,[Finanzas]" +
" ,[Finanzas_CierreProyecto]" +
" ,[Compras]" +
" ,[Dosier de Calidad]" +
" ,[Ventas]" +
" ,[Ventas_AltayDocumentosClientes]" +
" ,[Ventas_PedidoInterno]" +
" ,[Ventas_ExpedienteComercial]" +
" ,[Proyectos]" +
" ,[Proyectos_ExpedienteTecnico]" +
" ,[Proyectos_Planos]" +
" ,[EquipamientoeInstalaciones]" +
" ,[EquipamientoeInstalaciones_Fabricacion y Equipamiento]" +
" ,[EquipamientoeInstalaciones_Instalacionyarranque]" +
" ,[EquipamientoeInstalaciones_OperacionyMantenimiento]" +
" ,[General]" +


                                   "from [Proyectos_Avance] where Proyecto = @proyecto"

                                   , conexion2);
            SqlDataAdapter sda2 = new SqlDataAdapter(cmd2);
            cmd2.Parameters.AddWithValue("proyecto", Nombre);
            sda2.SelectCommand.CommandTimeout = 36000;
            DataTable dt2 = new DataTable();
            sda2.Fill(dt2);
            if (dt2.Rows.Count > 0)
            {
                AtencionClientes = dt2.Rows[0][0].ToString();
                AtencionClientes_Atencion = dt2.Rows[0][1].ToString();
                AtencionClientes_EntregaDocumentos = dt2.Rows[0][2].ToString();
                AtencionClientes_Postventa = dt2.Rows[0][3].ToString();
                AtencionClientes_Encuesta = dt2.Rows[0][4].ToString();
                Finanzas = dt2.Rows[0][5].ToString();
                Finanzas_CierreProyecto = dt2.Rows[0][6].ToString();
                Compraas = dt2.Rows[0][7].ToString();
                DosierdeCalidad = dt2.Rows[0][8].ToString();
                Ventas = dt2.Rows[0][9].ToString();
                Ventas_AltayDocumentosClientes = dt2.Rows[0][10].ToString();
                Ventas_PedidoInterno = dt2.Rows[0][11].ToString();
                Ventas_ExpedienteComercial = dt2.Rows[0][12].ToString();
                Proyectos = dt2.Rows[0][13].ToString();
                Proyectos_ExpedienteTecnic = dt2.Rows[0][14].ToString();
                Proyectos_Planos = dt2.Rows[0][15].ToString();
                EquipamientoeInstalaciones = dt2.Rows[0][16].ToString();
                EquipamientoeInstalaciones_FabricacionyEquipamiento = dt2.Rows[0][17].ToString();
                EquipamientoeInstalaciones_Instalacionyarranque = dt2.Rows[0][18].ToString();
                EquipamientoeInstalaciones_OperacionyMantenimiento = dt2.Rows[0][19].ToString();
                General = dt2.Rows[0][20].ToString();
                conexion2.Dispose();
           //     proyectos_Avance_INOXTableAdapter.UpdateTotalComun(General, Nombre);
            }
      
            else
            {

                proyectos_AvanceTableAdapter.AgregaProyecto(Nombre, "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0");
                proyectos_Avance_INOXTableAdapter.InsertQuery(Nombre = "0", "0", "0", "0", "0");
                proyectos_Avance_otherTableAdapter.InsertQuery(Nombre = "0", "0", "0", "0", "0","0","0","0");
                conexion2.Dispose();
            }
           
            if (Departamento == "Ventas")
            {
                int total = 11; decimal porcentaje;

                contador1 = contador1 + 1;
                porcentaje = (contador1 * 100) / 11;

                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                    Nombre,
                AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Porcentaje
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , General

            , Nombre);
            }
            else if (Departamento == "Atención a cliente")
            {
                int total = 25; decimal porcentaje;
                contador2 = contador2 + 1;
                porcentaje = (contador2 * 100) / 25;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(

                      Nombre,

                Porcentaje
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , General
                 , Nombre);

            }

            else if (Departamento == "PRODUCCION" || Departamento == "Produccion")
            {

                int total = 4; decimal porcentaje;
                contador3 = contador3 + 1;
                porcentaje = (contador3 * 100) / 4;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
                  AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , Porcentaje
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , General
                 , Nombre);

            }

            else if (Departamento == "FINANZAS")
            {
                int total = 4; decimal porcentaje;
                contador4 = contador4 + 1;
                porcentaje = (contador4 * 100) / 4;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
     AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Porcentaje
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , General
                 , Nombre);
            }

            else if (Departamento == "Proyectos")
            {
                int total = 8; decimal porcentaje;
                contador5 = contador5 + 1;
                porcentaje = (contador5 * 100) / 8;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
            AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Porcentaje
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , General
                 , Nombre);
            }

            else if (Departamento == "Equipamiento e Instalaciones")
            {
                int total = 23; decimal porcentaje;
                contador6 = contador6 + 1;
                porcentaje = (contador6 * 100) / 23;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
       AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , Porcentaje
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , General
                 , Nombre);
            }

            else if (Departamento == "Compras")
            {
                int total = 3; decimal porcentaje;
                contador7 = contador7 + 1;
                porcentaje = (contador7 * 100) / 3;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
        AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Porcentaje
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , General
                 , Nombre);
            }

            SqlConnection conexion3 = new SqlConnection(ObtenerCadena());
            conexion3.Open();
            SqlCommand cmd3 = new SqlCommand(
                                   "select " +
"  [AtencionClientes]" +
" ,[AtencionClientes_Atencion]" +
" ,[AtencionClientes_EntregaDocumentos]" +
" ,[AtencionClientes_Postventa]" +
" ,[AtencionClientes_Encuesta]" +
" ,[Finanzas]" +
" ,[Finanzas_CierreProyecto]" +
" ,[Compras]" +
" ,[Dosier de Calidad]" +
" ,[Ventas]" +
" ,[Ventas_AltayDocumentosClientes]" +
" ,[Ventas_PedidoInterno]" +
" ,[Ventas_ExpedienteComercial]" +
" ,[Proyectos]" +
" ,[Proyectos_ExpedienteTecnico]" +
" ,[Proyectos_Planos]" +
" ,[EquipamientoeInstalaciones]" +
" ,[EquipamientoeInstalaciones_Fabricacion y Equipamiento]" +
" ,[EquipamientoeInstalaciones_Instalacionyarranque]" +
" ,[EquipamientoeInstalaciones_OperacionyMantenimiento]" +
" ,[General]" +

                                   "from [Proyectos_Avance] where Proyecto = @proyecto"

                                   , conexion3);
            SqlDataAdapter sda3 = new SqlDataAdapter(cmd3);
            cmd3.Parameters.AddWithValue("proyecto", Nombre);
            sda3.SelectCommand.CommandTimeout = 36000;
            DataTable dt3 = new DataTable();
            sda3.Fill(dt3);
            if (dt3.Rows.Count > 0)
            {
                AtencionClientes = dt3.Rows[0][0].ToString();
                AtencionClientes_Atencion = dt3.Rows[0][1].ToString();
                AtencionClientes_EntregaDocumentos = dt3.Rows[0][2].ToString();
                AtencionClientes_Postventa = dt3.Rows[0][3].ToString();
                AtencionClientes_Encuesta = dt3.Rows[0][4].ToString();
                Finanzas = dt3.Rows[0][5].ToString();
                Finanzas_CierreProyecto = dt3.Rows[0][6].ToString();
                Compraas = dt3.Rows[0][7].ToString();
                DosierdeCalidad = dt3.Rows[0][8].ToString();
                Ventas = dt3.Rows[0][9].ToString();
                Ventas_AltayDocumentosClientes = dt3.Rows[0][10].ToString();
                Ventas_PedidoInterno = dt3.Rows[0][11].ToString();
                Ventas_ExpedienteComercial = dt3.Rows[0][12].ToString();
                Proyectos = dt3.Rows[0][13].ToString();
                Proyectos_ExpedienteTecnic = dt3.Rows[0][14].ToString();
                Proyectos_Planos = dt3.Rows[0][15].ToString();
                EquipamientoeInstalaciones = dt3.Rows[0][16].ToString();
                EquipamientoeInstalaciones_FabricacionyEquipamiento = dt3.Rows[0][17].ToString();
                EquipamientoeInstalaciones_Instalacionyarranque = dt3.Rows[0][18].ToString();
                EquipamientoeInstalaciones_OperacionyMantenimiento = dt3.Rows[0][19].ToString();
                General = dt3.Rows[0][20].ToString();
                conexion3.Dispose();
            }
            else { }

            SqlConnection conexion4 = new SqlConnection(ObtenerCadena());
            conexion4.Open();
            SqlCommand cmd4 = new SqlCommand(
                                   "select " +
"  [TotalComun]" +
" ,[ListadeMaterialesyEquipos]" +
" ,[InstalacionyPreparativos]" +
" ,[TotalGeneral]" +


                                   "from [Proyectos_Avance_INOX] where Proyecto = @proyecto"

                                   , conexion4);
            SqlDataAdapter sda4 = new SqlDataAdapter(cmd4);
            cmd4.Parameters.AddWithValue("proyecto", Nombre);
            sda4.SelectCommand.CommandTimeout = 36000;
            DataTable dt4 = new DataTable();
            sda4.Fill(dt4);
            if (dt4.Rows.Count > 0)
            {
                ListaInox = dt4.Rows[0][1].ToString();
                InstalacionInox = dt3.Rows[0][2].ToString();
                TotalComun = dt4.Rows[0][0].ToString();
                conexion4.Dispose();
            }
            else { }
            String ValorGeneral = General;
            Decimal Gen = Convert.ToDecimal(ValorGeneral);
            Decimal TotalGeneral = (Gen * 100) / 300;

            proyectos_Avance_INOXTableAdapter.UpdateGeneral(General, TotalGeneral.ToString(),Nombre);

            decimal ventas = Convert.ToDecimal(Ventas),
            Atencionacliente = Convert.ToDecimal(AtencionClientes),
            produccion = Convert.ToDecimal(EquipamientoeInstalaciones_FabricacionyEquipamiento),
            finanzas = Convert.ToDecimal(Finanzas),
            proyectos = Convert.ToDecimal(Proyectos), equip = Convert.ToDecimal(EquipamientoeInstalaciones);
            generalporcentaje = ventas + Atencionacliente + produccion + finanzas + proyectos + equip;
            porcentaje2 = (generalporcentaje * 100) / 600;
            string Porcentaje2 = porcentaje2.ToString();
            if (Carpeta == "Alta y Documentos Clientes")
            {
                int total = 6; decimal porcentaje;
                contador8 = contador8 + 1;
                porcentaje = (contador8 * 100) / 6;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
                AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Porcentaje
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , porcentaje2.ToString()
                 , Nombre


           );


            }
            else if (Carpeta == "Expediente Comercial")
            {
                int total = 4; decimal porcentaje;
                contador9 = contador9 + 1;
                porcentaje = (contador9 * 100) / 4;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
                 AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Porcentaje
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , porcentaje2.ToString()
                 , Nombre
           );
            }
            else if (Carpeta == "Pedido Interno")
            {
                int total = 1; decimal porcentaje;
                contador10 = contador10 + 1;
                porcentaje = (contador10 * 100) / 1;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
      AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Porcentaje
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , Porcentaje2
                 , Nombre
           );
            }
            else if (Carpeta == "Atencion Cliente")
            {

                decimal total = 7; decimal porcentaje;
                contador11 = contador11 + 1;
                porcentaje = (contador11 * 100) / 7;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
                  AtencionClientes
                , Porcentaje
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , Porcentaje2
                 , Nombre
           );

            }
            else if (Carpeta == "Encuestas de Satisfaccion")
            {

                decimal total = 1; decimal porcentaje;
                contador12 = contador12 + 1;
                porcentaje = (contador12 * 100) / 1;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
                AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , Porcentaje
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , Porcentaje2
                 , Nombre
           );
            }
            else if (Carpeta == "Entrega Documentos")
            {
                int total = 11; decimal porcentaje;
                contador13 = contador13 + 1;
                porcentaje = (contador13 * 100) / 11;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
                  AtencionClientes
                , AtencionClientes_Atencion
                , Porcentaje
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , Porcentaje2
                 , Nombre
           );
            }
            else if (Carpeta == "Postventa")
            {
                int total = 2; decimal porcentaje;
                contador14 = contador14 + 1;
                porcentaje = (contador14 * 100) / 2;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
            AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , Porcentaje
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , Porcentaje2
                 , Nombre
           );
            }
            else if (Carpeta == "Cierre Proyecto")
            {
                int total = 4; decimal porcentaje;
                contador15 = contador15 + 1;
                porcentaje = (contador15 * 100) / 4;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
           AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Porcentaje
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , Porcentaje2
                 , Nombre
           );
            }
            else if (Carpeta == "Expediente Tecnico")
            {
                int total = 4; decimal porcentaje;
                contador16 = contador16 + 1;
                porcentaje = (contador16 * 100) / 4;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
              AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Porcentaje
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , Porcentaje2
                 , Nombre
           );
            }
            else if (Carpeta == "Planos")
            {
                int total = 4; decimal porcentaje;
                contador17 = contador17 + 1;
                porcentaje = (contador17 * 100) / 4;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
                    AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Porcentaje
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , Porcentaje2
                 , Nombre
           );
            }
            else if (Carpeta == "Fabricacion y Equipamiento")
            {
                int total = 4; decimal porcentaje=0;
                contador18 = contador18 + 1;
                if (Tipo == "Proyectos-Inox") { porcentaje = (contador18 * 100) / 6; }
                else { porcentaje = (contador18 * 100) / 3; }



                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
                 AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , Porcentaje
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , Porcentaje2
                 , Nombre
           );
                
           
            }

            else if (Carpeta == "Instalacion y arranque" || Carpeta == "Instalacion y arranque/Reportes" || Carpeta == "Instalacion y arranque/Comprobaciones" || Carpeta == "Instalacion y arranque/Requisiciones" || Carpeta == "Instalacion y arranque/Dispersion" || Carpeta == "Instalacion y arranque/Almacen" || Carpeta == "Instalacion y arranque/Orden de Cambio" || Carpeta == "Instalacion y arranque/fotos")
            {
                int total = 6; decimal porcentaje;
                contador19 = contador19 + 1;
                porcentaje = (contador19 * 100) / 6;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
                 AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , Porcentaje
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , Porcentaje2
                 , Nombre
           );
                //    porcentaje2=Porcentaje2+porcentaje+
            }

            else if (Carpeta == "Instalacion Preparativos")
            {

                SqlConnection conexion5 = new SqlConnection(ObtenerCadena());
                conexion5.Open();
               
                SqlCommand cmd5 = new SqlCommand(
                                       "select " +
    "  [TotalComun]" +
    " ,[ListadeMaterialesyEquipos]" +
    " ,[InstalacionyPreparativos]" +
    " ,[TotalGeneral]" +


                                       "from [Proyectos_Avance_INOX] where Proyecto = @proyecto"

                                       , conexion5);
                SqlDataAdapter sda5 = new SqlDataAdapter(cmd5);
                cmd5.Parameters.AddWithValue("proyecto", Nombre);
                sda5.SelectCommand.CommandTimeout = 36000;
                DataTable dt5 = new DataTable();
                sda5.Fill(dt5);
                if (dt5.Rows.Count > 0)
                {
                    ListaInox = dt5.Rows[0][1].ToString();
                    InstalacionInox = dt5.Rows[0][2].ToString();
                    TotalComun = dt5.Rows[0][0].ToString();
                    conexion5.Dispose();
                }
                else { conexion5.Dispose(); }
                conexion5.Dispose();


                decimal porcentaje = 100;
                int total = 1;
                contador23 = contador23 + 1;
                porcentaje = (contador23 * 100) / 1;
                string Porcentaje = porcentaje.ToString();
                var convertDecimal = Convert.ToDecimal(Porcentaje2);
                var convertDecimal2 = Convert.ToDecimal(Porcentaje);
                var convertDecimal3 = Convert.ToDecimal(ListaInox);
                decimal porcentaje2;
                
                porcentaje2 = (convertDecimal + convertDecimal2 + convertDecimal3) / 3;
                     porcentaje = porcentaje / 1;
                string porcentajef = porcentaje2.ToString();
                proyectos_Avance_INOXTableAdapter.UpdatePreparativos(Porcentaje2, InstalacionInox, porcentaje2.ToString(), Nombre);

                //    contador27
                conexion5.Dispose();
            }

            else if (Carpeta == "Lista de Materiales y Equipos")
            {
                int total = 5; decimal porcentaje;
                contador27 = contador27 + 1;
                porcentaje = (contador27 * 100) / 7;
                string Porcentaje = porcentaje.ToString();

             
                var convertDecimal = Convert.ToDecimal(Porcentaje2);
                var convertDecimal2 = Convert.ToDecimal(Porcentaje);
                var convertDecimal3 = Convert.ToDecimal(ListaInox);
                decimal porcentaje2;
              
                porcentaje2 = (convertDecimal + convertDecimal2 + convertDecimal3) / 3;
                porcentaje = (porcentaje * 100) / 1;
                string porcentajef = porcentaje2.ToString();
                proyectos_Avance_INOXTableAdapter.UpdatePreparativos(Porcentaje2, ListaInox, porcentaje2.ToString(), Nombre);

                //    contador27
            }

            else if (Carpeta == "Obra Civil" || Carpeta == "Obra civil/Presupuesto/Ctg Obra" || Carpeta == "Obra civil/Subcontrato/Expediente de Obra" || Carpeta == "Obra civil/Subcontrato/Programa de Obra" || Carpeta == "Operacion y Mantenimiento/Dispersion" || Carpeta == "Instalacion y arranque/Almacen" || Carpeta == "Operacion y Mantenimiento/Orden de Cambio" || Carpeta == "Operacion y Mantenimiento/fotos")
            {
                int total = 8; decimal porcentaje;
                contador20 = contador20 + 1;
                porcentaje = (contador20 * 100) / 8;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
              AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , Porcentaje
                , Porcentaje2
                 , Nombre
           );
            }

            else if (Carpeta == "Operacion y Mantenimiento" || Carpeta == "Operacion y Mantenimiento/Reportes" || Carpeta == "Operacion y Mantenimiento/Comprobaciones" || Carpeta == "Operacion y Mantenimiento/Requisiciones" || Carpeta == "Operacion y Mantenimiento/Dispersion" || Carpeta == "Instalacion y arranque/Almacen" || Carpeta == "Operacion y Mantenimiento/Orden de Cambio" || Carpeta == "Operacion y Mantenimiento/fotos")
            {
                int total = 8; decimal porcentaje;
                contador20 = contador20 + 1;
                porcentaje = (contador20 * 100) / 8;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
              AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , Porcentaje
                , Porcentaje2
                 , Nombre
           );
            }

            else if (Carpeta == "Compras" || Carpeta == "Compras/Ordenes de Compras" || Carpeta == "Compras/Comparativas" || Carpeta == "Compras/Facturas")
            {
                int total = 3; decimal porcentaje;
                contador21 = contador21 + 1;
                porcentaje = (contador21 * 100) / 3;
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
             AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Porcentaje
                , DosierdeCalidad
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , Porcentaje2
                 , Nombre
           );
            }
            else if (Carpeta == "Dosier de Calidad" || Carpeta == "Dosier de Calidad/Certificados de calidad" || Carpeta == "Dosier de Calidad/Pruebas de Calidad/Estanquiedad" || Carpeta == "Dosier de Calidad/Pruebas de Calidad/Equipos" || Carpeta == "Dosier de Calidad/Pruebas de Calidad/Instalaciones" || Carpeta == "Dosier de Calidad/Pruebas de Calidad/Analisis Agua" || Carpeta == "Dosier de Calidad/Pruebas de Calidad/Carta Garantia" || Carpeta == "Dosier de Calidad/Pruebas de Calidad/Constancia" || Carpeta == "Dosier de Calidad/Pruebas de Calidad/Fotos" || Carpeta == "Dosier de Calidad/Pruebas de Calidad/Equipos" || Carpeta == "Dosier de Calidad/Certificado de Calidad" || Carpeta == "Dosier de Calidad/Revisiones")
            {

                decimal total = 23; decimal porcentaje = 0 ;
                contador22 = contador22 + 1;
                if (Tipo == "Proyectos-Inox") { porcentaje = (contador22 * 100) / 27; }
                else { porcentaje = (contador22 * 100) / 31; }
               
                string Porcentaje = porcentaje.ToString();
                proyectos_AvanceTableAdapter.Actualiza_index(
                      Nombre,
              AtencionClientes
                , AtencionClientes_Atencion
                , AtencionClientes_EntregaDocumentos
                , AtencionClientes_Postventa
                , AtencionClientes_Encuesta
                , Finanzas
                , Finanzas_CierreProyecto
                , Compraas
                , Porcentaje
                , Ventas
                , Ventas_AltayDocumentosClientes
                , Ventas_PedidoInterno
                , Ventas_ExpedienteComercial
                , Proyectos
                , Proyectos_ExpedienteTecnic
                , Proyectos_Planos
                , EquipamientoeInstalaciones
                , EquipamientoeInstalaciones_FabricacionyEquipamiento
                , EquipamientoeInstalaciones_Instalacionyarranque
                , EquipamientoeInstalaciones_OperacionyMantenimiento
                , Porcentaje2
                 , Nombre
           );
            }

            generalporcentaje = ventas + Atencionacliente + produccion + finanzas + proyectos + equip;
            porcentaje2 = (generalporcentaje * 100) / 600;
            proyectos_AvanceTableAdapter.ActualizaPorcentajefinal(porcentaje2.ToString(), Nombre);

        }

        public void ActualizaPorcentajesINOX()
        {

            SqlConnection conexion6 = new SqlConnection(ObtenerCadena());
            conexion6.Open();





            SqlCommand cmd6 = new SqlCommand(
                                   "select" +
                                     "[Departamento], " +
                                   "[Carpeta], " +
                                   "[Documento] " +

                                   "from [Ser_Documentos_INOX] where Documento = @Nombre"

                                   , conexion6);
            SqlDataAdapter sda6 = new SqlDataAdapter(cmd6);
            cmd6.Parameters.AddWithValue("Nombre", Path.GetFileName(files[i]));
            sda6.SelectCommand.CommandTimeout = 36000;
            DataTable dt6 = new DataTable();
            sda6.Fill(dt6);
            conexion6.Dispose();
            if (dt6.Rows.Count > 0)
            {
                Documento = dt6.Rows[0][2].ToString();
                Carpeta = dt6.Rows[0][1].ToString();
                Departamento = dt6.Rows[0][0].ToString();
                Thread.Sleep(50);


            }



            if (Carpeta == "Instalacion Preparativos")
            {

                SqlConnection conexion = new SqlConnection(ObtenerCadena());
                conexion.Open();
                SqlCommand cmd = new SqlCommand(
                                       "select" +
                                       "[Ano], " +
                                       "[tipo2], " +
                                       "[Nombre], " +
                                       "[Nombre2] " +

                                       "from [Folio_proyectos>2019] where Nombre = @Nombre"

                                       , conexion);
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                cmd.Parameters.AddWithValue("Nombre", Folio);
                sda.SelectCommand.CommandTimeout = 36000;
                DataTable dt = new DataTable();
                sda.Fill(dt);
                if (dt.Rows.Count > 0)
                {

                    NombreProyecto = dt.Rows[0][3].ToString();
                    Thread.Sleep(50);

                    conexion.Dispose();
                }

                conexion.Dispose();

                SqlConnection conexion4 = new SqlConnection(ObtenerCadena());
                conexion4.Open();





                SqlCommand cmd4 = new SqlCommand(
                                       "select" +
                                         "[Proyecto], " +
                                       "[ListadeMaterialesyEquipos], " +
                                       "[InstalacionyPreparativos]" +

                                       "from [Proyectos_Avance_INOX] where Proyecto = @Nombre"

                                       , conexion4);
                SqlDataAdapter sda4 = new SqlDataAdapter(cmd4);
                cmd4.Parameters.AddWithValue("Nombre", NombreProyecto);
                sda4.SelectCommand.CommandTimeout = 36000;
                DataTable dt4 = new DataTable();
                sda4.Fill(dt4);
                conexion4.Dispose();
                if (dt4.Rows.Count > 0)
                {
                    ValorPreparativos = dt4.Rows[0][2].ToString();
                    ValorLista = dt4.Rows[0][1].ToString();

                    Thread.Sleep(50);

                    if (ValorLista == "" || ValorLista == null) { ValorLista = "0"; }
                    if (ValorPreparativos == "" || ValorPreparativos == null) { ValorPreparativos = "0"; }
                }
                else {
                    string PorcentajeComun = porcentaje2.ToString(); 
                    proyectos_Avance_INOXTableAdapter.InsertQuery(NombreProyecto, PorcentajeComun, "0", "0", "0"); }
                int total = 11; decimal porcentaje, porcentaje3, porcentaje4;
                
                string Porcentaje2 = porcentaje2.ToString();
                contador25 = contador25 + 1;
                porcentaje = (contador25 * 100) / 1;
                if (ValorLista == "" || ValorLista == null || ValorLista == " ") { ValorLista = "0"; }
                decimal Lista = Convert.ToDecimal(ValorLista);
                if (ValorPreparativos == "" || ValorPreparativos == null || ValorPreparativos == " ") { ValorPreparativos = "0"; }

                decimal Preparativos = Convert.ToDecimal(ValorPreparativos);
                //Porcentaje de carpeta
                string Porcentaje = porcentaje.ToString();
                decimal Totalpreparativos= Preparativos + porcentaje;
                if (Totalpreparativos > 100) { }
                else if (Totalpreparativos <= 100)
                {
                    porcentaje3 = porcentaje2 + Lista + Totalpreparativos;
                    porcentaje4 = porcentaje3 * 100 / 300;
                    string PorcentajeFinal = porcentaje4.ToString();
                    proyectos_Avance_INOXTableAdapter.UpdatePreparativos(Porcentaje2, Porcentaje, PorcentajeFinal, NombreProyecto);
                }

              
               

            }
            else if (Carpeta == "Lista de Materiales y Equipos")
            {
                SqlConnection conexion = new SqlConnection(ObtenerCadena());
                conexion.Open();
                SqlCommand cmd = new SqlCommand(
                                       "select" +
                                       "[Ano], " +
                                       "[tipo2], " +
                                       "[Nombre], " +
                                       "[Nombre2] " +

                                       "from [Folio_proyectos>2019] where Nombre = @Nombre"

                                       , conexion);
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                cmd.Parameters.AddWithValue("Nombre", Folio);
                sda.SelectCommand.CommandTimeout = 36000;
                DataTable dt = new DataTable();
                sda.Fill(dt);
                if (dt.Rows.Count > 0)
                {

                    NombreProyecto = dt.Rows[0][3].ToString();
                    Thread.Sleep(50);

                    conexion.Dispose();
                }
                else { proyectos_Avance_INOXTableAdapter.InsertQuery(NombreProyecto, "0", "0", "0", "0"); }

                conexion.Dispose();



                String valor = "", Totalfinal = "", ValorFabricacion = "", ValorPreparativos = "";
                SqlConnection conexion4 = new SqlConnection(ObtenerCadena());
                conexion4.Open();





                SqlCommand cmd4 = new SqlCommand(
                                       "select" +
                                         "[Proyecto], " +
                                       "[ListadeMaterialesyEquipos], " +
                                       "[InstalacionyPreparativos]" +

                                       "from [Proyectos_Avance_INOX] where Proyecto = @Nombre"

                                       , conexion4);
                SqlDataAdapter sda4 = new SqlDataAdapter(cmd4);
                cmd4.Parameters.AddWithValue("Nombre", NombreProyecto);
                sda4.SelectCommand.CommandTimeout = 36000;
                DataTable dt4 = new DataTable();
                sda4.Fill(dt4);
                conexion4.Dispose();
                if (dt4.Rows.Count > 0)
                {
                    ValorPreparativos = dt4.Rows[0][2].ToString();
                    ValorFabricacion = dt4.Rows[0][1].ToString();

                    Thread.Sleep(50);


                }
                else { proyectos_Avance_INOXTableAdapter.InsertQuery(NombreProyecto, "", "0", "0", "0"); }
                string Porcentaje2 = porcentaje2.ToString();
                int total = 11; decimal porcentaje, porcentaje3, porcentaje4;

                contador26 = contador26 + 1;


       

                /// 5 archivos
                /// 
                porcentaje = (contador25 * 100) / 5;
                decimal Preparativos = Convert.ToDecimal(ValorPreparativos);
                string Porcentaje = porcentaje.ToString();
                decimal Lista = Convert.ToDecimal(ListaInox);

                porcentaje3 = porcentaje2 + Lista + Preparativos;


                porcentaje4 = (porcentaje3 * 100) / 300;
                string PorcentajeFinal = porcentaje4.ToString();
                proyectos_Avance_INOXTableAdapter.UpdateFabricacion(Porcentaje2, Porcentaje, PorcentajeFinal, NombreProyecto);

                SqlConnection conexion5 = new SqlConnection(ObtenerCadena());
                conexion5.Open();

                SqlCommand cmd5 = new SqlCommand(
                                       "select " +
    "  [TotalComun]" +
    " ,[ListadeMaterialesyEquipos]" +
    " ,[InstalacionyPreparativos]" +
    " ,[TotalGeneral]" +


                                       "from [Proyectos_Avance_INOX] where Proyecto = @proyecto"

                                       , conexion5);
                SqlDataAdapter sda5 = new SqlDataAdapter(cmd5);
                cmd5.Parameters.AddWithValue("proyecto", Nombre);
                sda5.SelectCommand.CommandTimeout = 36000;
                DataTable dt5 = new DataTable();
                sda5.Fill(dt5);
                if (dt5.Rows.Count > 0)
                {
                    ListaInox = dt5.Rows[0][1].ToString();
                    InstalacionInox = dt5.Rows[0][2].ToString();
                    TotalComun = dt5.Rows[0][0].ToString();
                    conexion5.Dispose();
                }
                else { conexion5.Dispose(); }
                conexion5.Dispose();
                 Lista = Convert.ToDecimal(ListaInox);

                porcentaje3 = porcentaje2 + Lista + Preparativos;

                porcentaje4 = (porcentaje3 * 100) / 300;
                 PorcentajeFinal = porcentaje4.ToString();

                proyectos_Avance_INOXTableAdapter.UpdateFabricacion(Porcentaje2, Porcentaje, PorcentajeFinal, NombreProyecto);
            }
            else{
                if (ValorLista == "" || ValorLista == null) { ValorLista = "0"; }
                if (ListaInox == "" || ListaInox == null) { ListaInox = "0"; }
                if (ValorPreparativos == "" || ValorPreparativos == null) { ValorPreparativos = "0"; }
                if (ValorLista == "0" && ValorPreparativos == "0")
                {
                    decimal Preparativos = Convert.ToDecimal(ValorPreparativos);
                    decimal Comun = Convert.ToDecimal(porcentaje2);
                    decimal Lista = Convert.ToDecimal(ListaInox);

                    decimal sumatotal = Lista + Preparativos + Comun;
                    decimal totalgeneral = sumatotal * 100 / 300;
                    string Porcentaje2 = porcentaje2.ToString();
                    string TotalGeneral = totalgeneral.ToString();
                    proyectos_Avance_INOXTableAdapter.UpdateTotales(Porcentaje2, TotalGeneral, NombreProyecto);
                }
                else { }
              
            }
        }
        public void ActualizaPorcentajesINOXFinal()
        {

          



                SqlConnection conexion4 = new SqlConnection(ObtenerCadena());
                conexion4.Open();





                SqlCommand cmd4 = new SqlCommand(
                                       "select" +
                                         "[Proyecto], " +
                                       "[ListadeMaterialesyEquipos], " +
                                       "[InstalacionyPreparativos]" +

                                       "from [Proyectos_Avance_INOX] where Proyecto = @Nombre"

                                       , conexion4);
                SqlDataAdapter sda4 = new SqlDataAdapter(cmd4);
                cmd4.Parameters.AddWithValue("Nombre", NombreProyecto);
                sda4.SelectCommand.CommandTimeout = 36000;
                DataTable dt4 = new DataTable();
                sda4.Fill(dt4);
                conexion4.Dispose();
                if (dt4.Rows.Count > 0)
                {
                    ValorPreparativos = dt4.Rows[0][2].ToString();
                    ValorLista = dt4.Rows[0][1].ToString();

                    Thread.Sleep(50);

                   
                }
                else
                {
                  
                }
            if (ValorLista == "" || ValorLista == null) { ValorLista = "0"; }
            if (ValorPreparativos == "" || ValorPreparativos == null) { ValorPreparativos = "0"; }

            decimal Preparativos = Convert.ToDecimal(ValorPreparativos);
                decimal Comun = Convert.ToDecimal(porcentaje2);
                decimal Lista = Convert.ToDecimal(ListaInox);

                decimal sumatotal = Lista + Preparativos + Comun;
                decimal totalgeneral = sumatotal * 100 / 300;
                string Porcentaje2 = porcentaje2.ToString();
                string TotalGeneral = totalgeneral.ToString();
                proyectos_Avance_INOXTableAdapter.UpdateTotales(Porcentaje2, TotalGeneral, NombreProyecto);

            

         
        }

        public void ActualizaPorcentajesINDPot()
        {

            SqlConnection conexion6 = new SqlConnection(ObtenerCadena());
            conexion6.Open();





            SqlCommand cmd6 = new SqlCommand(
                                   "select" +
                                     "[Departamento], " +
                                   "[Carpeta], " +
                                   "[Documento] " +

                                   "from [Ser_Documentos_IND-POT] where Documento = @Nombre"

                                   , conexion6);
            SqlDataAdapter sda6 = new SqlDataAdapter(cmd6);
            cmd6.Parameters.AddWithValue("Nombre", Path.GetFileName(files[i]));
            sda6.SelectCommand.CommandTimeout = 36000;
            DataTable dt6 = new DataTable();
            sda6.Fill(dt6);
            conexion6.Dispose();
            if (dt6.Rows.Count > 0)
            {
                Documento = dt6.Rows[0][2].ToString();
                Carpeta = dt6.Rows[0][1].ToString();
                Departamento = dt6.Rows[0][0].ToString();
                Thread.Sleep(50);


            }



            if (Carpeta == "Obra civil/Presupuesto/Ctg Obra" 
                || Carpeta == "Obra civil/Subcontrato/Expediente de Obra"
                || Carpeta == "Obra civil/Subcontrato/Programa de Obra"
                || Carpeta == "Obra civil/Subcontrato/Consolidados"
                || Carpeta == "Obra civil/Control de Obra/Liberaciones"
                || Carpeta == "Obra civil/Control de Obra/Requisiciones"
                || Carpeta == "Obra civil/Control de Obra/Bitacora"
                || Carpeta == "Obra civil/Control de Obra/Dispersiones"
                || Carpeta == "Obra civil/Control de Obra/Almacen"
                || Carpeta == "Obra civil/Control de Obra/Comprobacion Gastos"
                || Carpeta == "Obra civil/Control de Obra/Asistencia"
                || Carpeta == "Obra civil/Control de Obra/Reporte/Fotografico"
                || Carpeta == "Obra civil/Control de Obra/Reporte/Fisico Fin"
                || Carpeta == "Obra civil/Control de Obra/Reportes"
                || Carpeta == "Obra civil/Control de Obra/Estimaciones"
                || Carpeta == "Obra civil/Control de Obra/Reportes/Direccion"
                || Carpeta == "Obra civil/Control de Obra/Dosier de Calidad"
                || Carpeta == "Obra civil/Control de Obra/Orden de Cambios"
                || Carpeta == "Obra civil/Control de Obra/Seguridad e Higiene"
                || Carpeta == "Obra civil/Control de Obra/fotos"
                || Carpeta == "Obra civil/Control de Obra/Lista de Materiales y equipos"
                || Carpeta == "Obra civil/SIROC E IMSS/Registro de Siroc"
                || Carpeta == "Obra civil/SIROC E IMSS/Subcontratistas"
                || Carpeta == "Obra civil/SIROC E IMSS/Altas"
                || Carpeta == "Obra civil/SIROC E IMSS/Pagos Imss"
                || Carpeta == "Obra civil/Subcontratos/Contrato"
                || Carpeta == "Obra civil/Subcontratos/Presupuesto"
                || Carpeta == "Obra civil/Subcontratos/Concentrado Estimaciones"
                || Carpeta == "Obra civil/Subcontratos/Estimaciones"  )
            {

                SqlConnection conexion = new SqlConnection(ObtenerCadena());
                conexion.Open();
                SqlCommand cmd = new SqlCommand(
                                       "select" +
                                       "[Ano], " +
                                       "[tipo2], " +
                                       "[Nombre], " +
                                       "[Nombre2] " +

                                       "from [Folio_proyectos>2019] where Nombre = @Nombre"

                                       , conexion);
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                cmd.Parameters.AddWithValue("Nombre", Folio);
                sda.SelectCommand.CommandTimeout = 36000;
                DataTable dt = new DataTable();
                sda.Fill(dt);
                if (dt.Rows.Count > 0)
                {

                    NombreProyecto = dt.Rows[0][3].ToString();
                    Thread.Sleep(50);

                    conexion.Dispose();
                }

                conexion.Dispose();

                SqlConnection conexion4 = new SqlConnection(ObtenerCadena());
                conexion4.Open();





                SqlCommand cmd4 = new SqlCommand(
                                       "select" +
                                         "[Proyecto], " +
                                       "[Obracivil], " +
                                       "[SiroceImss]," +
                                         "[Subcontratos]," +
                                         "[ControldeObra]," +
                                         "[Presupuesto]" +


                                       "from [Proyectos_Avance_other] where Proyecto = @Nombre"

                                       , conexion4);
                SqlDataAdapter sda4 = new SqlDataAdapter(cmd4);
                cmd4.Parameters.AddWithValue("Nombre", NombreProyecto);
                sda4.SelectCommand.CommandTimeout = 36000;
                DataTable dt4 = new DataTable();
                sda4.Fill(dt4);
                conexion4.Dispose();
                if (dt4.Rows.Count > 0)
                {
                    ObraCivil = dt4.Rows[0][1].ToString();
                    SirocEIMSS = dt4.Rows[0][2].ToString();
                    Subcontratos = dt4.Rows[0][3].ToString();
                    ControldeObra = dt4.Rows[0][4].ToString();
                    Presupuesto = dt4.Rows[0][5].ToString();

                    Thread.Sleep(50);

                    if (ObraCivil == "" || ObraCivil == null) { ObraCivil = "0"; }
                    if (SirocEIMSS == "" || SirocEIMSS == null) { SirocEIMSS = "0"; }
                    if (Subcontratos == "" || Subcontratos == null) { Subcontratos = "0"; }
                    if (ControldeObra == "" || ControldeObra == null) { ControldeObra = "0"; }
                    if (Presupuesto == "" || Presupuesto == null) { Presupuesto = "0"; }

                }
                else {
                    string PorcentajeComun = porcentaje2.ToString();
                    proyectos_Avance_otherTableAdapter.InsertQuery(NombreProyecto, PorcentajeComun, "0", "0", "0","0","0","0"); }
                int total = 47; decimal porcentaje, SumatoriaGeneral, Porcentajefinal;

              

                if (Carpeta == "Obra civil/Subcontratos/Contrato" || Carpeta == "Obra civil/Subcontratos/Presupuesto" || Carpeta == "Obra civil/Subcontratos/Concentrado Estimaciones" || Carpeta == "Obra civil/Subcontratos/Estimaciones"
                || Carpeta == "Obra civil/Subcontrato/Expediente de Obra" || Carpeta == "Obra civil/Subcontrato/Programa de Obra" || Carpeta == "Obra civil/Subcontrato/Consolidados")
                {
                    string Porcentaje2 = porcentaje2.ToString();
                    contador24 = contador24 + 1;
                    porcentaje = (contador24 * 100) / 19;

                    if (Subcontratos == "" || Subcontratos == null || Subcontratos == " ") { Subcontratos = "0"; }
                    decimal Valor = Convert.ToDecimal(Subcontratos);
           
            
                    //Porcentaje de carpeta
                    string Porcentaje = porcentaje.ToString();
                    decimal ObraCivild = Convert.ToDecimal(ObraCivil); 
                    decimal SirocEIMSSd = Convert.ToDecimal(SirocEIMSS); 
                    decimal Subcontratosd = Convert.ToDecimal(Subcontratos); 
                    decimal ControldeObrad = Convert.ToDecimal(ControldeObra); 
                    decimal Presupuestod = Convert.ToDecimal(Presupuesto);
                    decimal Totalindus = ObraCivild + SirocEIMSSd + ControldeObrad + Presupuestod+Valor;
                    SumatoriaGeneral = porcentaje2 + Totalindus;
                    Porcentajefinal = (SumatoriaGeneral * 100) / 600;
                        string PorcentajeFinal = Porcentajefinal.ToString();
                        proyectos_Avance_otherTableAdapter.UpdateSubContratos(NombreProyecto,Porcentaje2, Porcentaje, PorcentajeFinal);
                    

                }
                else if (Carpeta == "Obra civil/Presupuesto/Ctg Obra")
                {
                    string Porcentaje2 = porcentaje2.ToString();
                    contador28 = contador28 + 1; ///////////////////////////////////////////cambiar
                    porcentaje = (contador28 * 100) / 6;

                    if (Presupuesto == "" || Presupuesto == null || Presupuesto == " ") { Presupuesto = "0"; }
                    decimal Valor = Convert.ToDecimal(Presupuesto);


                    //Porcentaje de carpeta
                    string Porcentaje = porcentaje.ToString();
                    decimal ObraCivild = Convert.ToDecimal(ObraCivil);
                    decimal SirocEIMSSd = Convert.ToDecimal(SirocEIMSS);
                    decimal Subcontratosd = Convert.ToDecimal(Subcontratos);
                    decimal ControldeObrad = Convert.ToDecimal(ControldeObra);
                    decimal Presupuestod = Convert.ToDecimal(Presupuesto);
                    decimal Totalindus = ObraCivild + SirocEIMSSd + ControldeObrad + Subcontratosd + Valor;
                    SumatoriaGeneral = porcentaje2 + Totalindus;
                    Porcentajefinal = (SumatoriaGeneral * 100) / 600;
                    string PorcentajeFinal = Porcentajefinal.ToString();
                    proyectos_Avance_otherTableAdapter.UpdatePresupuesto(NombreProyecto, Porcentaje2, Porcentaje, PorcentajeFinal);

                }
                else if (Carpeta == "Obra civil/Control de Obra/Liberaciones" || Carpeta == "Obra civil/Control de Obra/Requisiciones" || Carpeta == "Obra civil/Control de Obra/Bitacora" || Carpeta == "Obra civil/Control de Obra/Dispersiones"
                || Carpeta == "Obra civil/Control de Obra/Almacen" || Carpeta == "Obra civil/Control de Obra/Comprobacion Gastos" || Carpeta == "Obra civil/Control de Obra/Asistencia" || Carpeta == "Obra civil/Control de Obra/Reporte/Fotografico" || Carpeta == "Obra civil/Control de Obra/Reporte/Fisico Fin" || Carpeta == "Obra civil/Control de Obra/Reportes"
                || Carpeta == "Obra civil/Control de Obra/Estimaciones" || Carpeta == "Obra civil/Control de Obra/Reportes/Direccion" || Carpeta == "Obra civil/Control de Obra/Dosier de Calidad" || Carpeta == "Obra civil/Control de Obra/Orden de Cambios"
                || Carpeta == "Obra civil/Control de Obra/Seguridad e Higiene" || Carpeta == "Obra civil/Control de Obra/fotos" || Carpeta == "Obra civil/Control de Obra/Lista de Materiales y equipos")
                {
                    string Porcentaje2 = porcentaje2.ToString();
                    contador29 = contador29 + 1;///////////////////////////////////
                    porcentaje = (contador29 * 100) / 18;

                    if (ControldeObra == "" || ControldeObra == null || ControldeObra == " ") { ControldeObra = "0"; }
                    decimal Valor = Convert.ToDecimal(ControldeObra);


                    //Porcentaje de carpeta
                    string Porcentaje = porcentaje.ToString();
                    decimal ObraCivild = Convert.ToDecimal(ObraCivil);
                    decimal SirocEIMSSd = Convert.ToDecimal(SirocEIMSS);
                    decimal Subcontratosd = Convert.ToDecimal(Subcontratos);
                    decimal ControldeObrad = Convert.ToDecimal(ControldeObra);
                    decimal Presupuestod = Convert.ToDecimal(Presupuesto);
                    decimal Totalindus = ObraCivild + SirocEIMSSd + Presupuestod + Subcontratosd + Valor;
                    SumatoriaGeneral = porcentaje2 + Totalindus;
                    Porcentajefinal = (SumatoriaGeneral * 100) / 600;
                    string PorcentajeFinal = Porcentajefinal.ToString();
                    proyectos_Avance_otherTableAdapter.UpdateControldeObra(NombreProyecto, Porcentaje2, Porcentaje, PorcentajeFinal);


                }
                else if (Carpeta == "Obra civil/SIROC E IMSS/Registro de Siroc" || Carpeta == "Obra civil/SIROC E IMSS/Subcontratistas" || Carpeta == "Obra civil/SIROC E IMSS/Altas" || Carpeta == "Obra civil/SIROC E IMSS/Pagos Imss")
                {
                    string Porcentaje2 = porcentaje2.ToString();
                    contador30 = contador30 + 1; ///////////////////////////cambiar
                    porcentaje = (contador30 * 100) / 4;

                    if (SirocEIMSS == "" || SirocEIMSS == null || SirocEIMSS == " ") { SirocEIMSS = "0"; }
                    decimal Valor = Convert.ToDecimal(SirocEIMSS);


                    //Porcentaje de carpeta
                    string Porcentaje = porcentaje.ToString();
                    decimal ObraCivild = Convert.ToDecimal(ObraCivil);
                    decimal SirocEIMSSd = Convert.ToDecimal(SirocEIMSS);
                    decimal Subcontratosd = Convert.ToDecimal(Subcontratos);
                    decimal ControldeObrad = Convert.ToDecimal(ControldeObra);
                    decimal Presupuestod = Convert.ToDecimal(Presupuesto);
                    decimal Totalindus = ObraCivild + Subcontratosd + ControldeObrad + Presupuestod + Valor;
                    SumatoriaGeneral = porcentaje2 + Totalindus;
                    Porcentajefinal = SumatoriaGeneral * 100 / 600;
                    string PorcentajeFinal = Porcentajefinal.ToString();
                    proyectos_Avance_otherTableAdapter.UpdateSiroc(NombreProyecto, Porcentaje2, Porcentaje, PorcentajeFinal);



                }

                ///////////////////////validador si esta en ceros el proyecto manda almenos el comun




            }

            else

             
            {
                if (ObraCivil == "" || ObraCivil == null) { ObraCivil = "0"; }
                if (SirocEIMSS == "" || SirocEIMSS == null) { SirocEIMSS = "0"; }
                if (Subcontratos == "" || Subcontratos == null) { Subcontratos = "0"; }
                if (ControldeObra == "" || ControldeObra == null) { ControldeObra = "0"; }
                if (Presupuesto == "" || Presupuesto == null) { Presupuesto = "0"; }

                if (ObraCivil == "0" && SirocEIMSS == "0" && Subcontratos == "0" && ControldeObra == "0" && Presupuesto == "0")
                {
                    decimal ObraCivild = Convert.ToDecimal(ObraCivil);
                    decimal SirocEIMSSd = Convert.ToDecimal(SirocEIMSS);
                    decimal Subcontratosd = Convert.ToDecimal(Subcontratos);
                    decimal ControldeObrad = Convert.ToDecimal(ControldeObra);
                    decimal Presupuestod = Convert.ToDecimal(Presupuesto);

                    decimal sumatotal = ObraCivild + SirocEIMSSd + Subcontratosd + ControldeObrad + Presupuestod + porcentaje2;
                    decimal totalgeneral = (sumatotal * 100) / 600;
                    string Porcentaje2 = porcentaje2.ToString();
                    string TotalGeneral = totalgeneral.ToString();
                    proyectos_Avance_otherTableAdapter.UpdateTotal(Porcentaje2, TotalGeneral, NombreProyecto);
                }
                else {
                    decimal ObraCivild = Convert.ToDecimal(ObraCivil);
                    decimal SirocEIMSSd = Convert.ToDecimal(SirocEIMSS);
                    decimal Subcontratosd = Convert.ToDecimal(Subcontratos);
                    decimal ControldeObrad = Convert.ToDecimal(ControldeObra);
                    decimal Presupuestod = Convert.ToDecimal(Presupuesto);

                    decimal sumatotal = ObraCivild + SirocEIMSSd + Subcontratosd + ControldeObrad + Presupuestod + porcentaje2;
                    decimal totalgeneral = (sumatotal * 100) / 600;
                    string Porcentaje2 = porcentaje2.ToString();
                    string TotalGeneral = totalgeneral.ToString();
                    proyectos_Avance_otherTableAdapter.UpdateTotal(Porcentaje2, TotalGeneral, NombreProyecto);
                }

            }




        }
        public void ActualizaPorcentajesINDPotFinal()
        {

           
                SqlConnection conexion4 = new SqlConnection(ObtenerCadena());
                conexion4.Open();





                SqlCommand cmd4 = new SqlCommand(
                                       "select" +
                                         "[Proyecto], " +
                                       "[Obracivil], " +
                                       "[SiroceImss]," +
                                         "[Subcontratos]," +
                                         "[ControldeObra]," +
                                         "[Presupuesto]" +


                                       "from [Proyectos_Avance_other] where Proyecto = @Nombre"

                                       , conexion4);
                SqlDataAdapter sda4 = new SqlDataAdapter(cmd4);
                cmd4.Parameters.AddWithValue("Nombre", NombreProyecto);
                sda4.SelectCommand.CommandTimeout = 36000;
                DataTable dt4 = new DataTable();
                sda4.Fill(dt4);
                conexion4.Dispose();
                if (dt4.Rows.Count > 0)
                {
                    ObraCivil = dt4.Rows[0][1].ToString();
                    SirocEIMSS = dt4.Rows[0][2].ToString();
                    Subcontratos = dt4.Rows[0][3].ToString();
                    ControldeObra = dt4.Rows[0][4].ToString();
                    Presupuesto = dt4.Rows[0][5].ToString();

                    Thread.Sleep(50);


                }
                else
                {
                  
                }
                int total = 47; decimal porcentaje, SumatoriaGeneral, Porcentajefinal;




                if (ObraCivil == "" || ObraCivil == null) { ObraCivil = "0"; }
                if (SirocEIMSS == "" || SirocEIMSS == null) { SirocEIMSS = "0"; }
                if (Subcontratos == "" || Subcontratos == null) { Subcontratos = "0"; }
                if (ControldeObra == "" || ControldeObra == null) { ControldeObra = "0"; }
                if (Presupuesto == "" || Presupuesto == null) { Presupuesto = "0"; }

             
                    decimal ObraCivild = Convert.ToDecimal(ObraCivil);
                    decimal SirocEIMSSd = Convert.ToDecimal(SirocEIMSS);
                    decimal Subcontratosd = Convert.ToDecimal(Subcontratos);
                    decimal ControldeObrad = Convert.ToDecimal(ControldeObra);
                    decimal Presupuestod = Convert.ToDecimal(Presupuesto);

                    decimal sumatotal = ObraCivild + SirocEIMSSd + Subcontratosd + ControldeObrad + Presupuestod + porcentaje2;
                    decimal totalgeneral = (sumatotal * 100) / 600;
                    string Porcentaje2 = porcentaje2.ToString();
                    string TotalGeneral = totalgeneral.ToString();
                    proyectos_Avance_otherTableAdapter.UpdateTotal(Porcentaje2, TotalGeneral, NombreProyecto);
                

            }




        }

    }


