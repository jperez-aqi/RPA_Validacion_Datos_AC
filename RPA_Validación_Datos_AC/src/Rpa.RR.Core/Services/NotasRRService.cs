using AutoHotkey.Interop;
using AutoHotkey.Interop.Pipes;
using AutoHotkey.Interop.Util;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Tokens;
using RPA_Validación_Datos_AC.src.Rpa.RR.Core.Utilities;
using RPA_Validación_Datos_AC.src.Rpa.RR.Infrastructure.Repositories;
using RPAExtraccionNotasRR.src.Rpa.RR.Core.Entities;
using RPAExtraccionNotasRR.src.Rpa.RR.Core.Utilities;
using RPAExtraccionNotasRR.src.Rpa.RR.Helpers;
using RPAExtraccionNotasRR.src.Rpa.RR.Infrastructure.Email;
using RPAExtraccionNotasRR.src.Rpa.RR.Infrastructure.Repositories;
using System;
using System.Buffers.Text;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace RPAExtraccionNotasRR.src.Rpa.RR.Core.Services
{
    public class NotasRRService : INotasRRService
    {
        private readonly CasoRepository _casoRepo;
        private readonly CredencialesRepository _credencialesRepo;
        private readonly ActualizarCasoRepository _actualizarCasoRepo;
        private readonly IEmailSender _emailSender;
        private readonly string _clave;
        private readonly int _maxLoginAttempts;
        public static Logtext log = new Logtext();
        private static string usuario;
        private static string clave;
        


        public NotasRRService(
            CasoRepository casoRepo,
            CredencialesRepository credencialesRepo,
            ActualizarCasoRepository actualizarCasoRepo,
            IEmailSender emailSender,
            IConfiguration config
            ) 
        {
            _casoRepo = casoRepo;
            _credencialesRepo = credencialesRepo;
            _actualizarCasoRepo = actualizarCasoRepo;
            _emailSender = emailSender;
            _clave = config["Ac:Clave"]; 
            _maxLoginAttempts = config.GetValue<int>("RetryPolicy:MaxAttempts", 3);
        }

        public async Task Run()
        {
            log.LogInit();
            log.Escribir("Iniciando RPA de Validación Datos AC...");

            _casoRepo.ResetFallidos(); // Reiniciar casos fallidos
            //_actualizarCasoRepo.InsertarDetalle(618731);
            //const int APLICATIVO_ID = 3024;  // Pon aquí el Id que corresponda
            // 2) Obtener credenciales de la base de datos
            log.Escribir("Obteniendo credenciales de la base de datos...");
            /*var creds = _credencialesRepo.ObtenerCredenciales(APLICATIVO_ID);

            usuario = creds.FirstOrDefault(c => c.Clave == "Usuario")?.Valor
                          ?? throw new Exception("No se encontró clave 'Usuario'");
            clave = creds.FirstOrDefault(c => c.Clave == "Clave")?.Valor
                          ?? throw new Exception("No se encontró clave 'Clave'");*/

            clave = _clave;

            ACHelperService ACHelper = new ACHelperService();

            CasoDto caso = new CasoDto();

            var ahk = AutoHotkeyEngine.Instance; 
            string RutaImagenGlobal = @"C:\Desarrollos-BI\ImagenesRobots\";
            string status = string.Empty;
            string custcode = string.Empty; 
            string nombre = string.Empty;
            string apellido = string.Empty;
            string nombreCompleto = string.Empty;
            string identificacion = string.Empty;
            string correo = string.Empty;
            string celular = string.Empty;
            bool BusquedaCuscode = false;

            // Lista de remitentes que deben ir a atención manual (todo en minúsculas, sin espacios)
            var blockedSenders = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "solucionespqr@conexionescelulares.com",
                "conexpqrleticia@conexionescelulares.com",
                "cpsarauca@datosmoviles.net",
                "1servicioalclientepqr@gmail.com",
                "barracpsinirida@movilco.com.co",
                "barracvssanjose@movilco.com.co",
                "barracpscarreno@movilco.com.co",
                "coordinadorapuertoasis@jesmarcomunicaciones.com",
                "contactenos@sic.gov.co",
                "notificacionesclaro@claro.com.co",
                "pqrstrasladogobierno@claro.com.co",
                "comunicacioncrccare@crcom.gov.co",
                "radicador2@claro.com.co",
                "cpsinirida@movilco.co"
            };
            string remitente = string.Empty;
            string remitenteNorm = string.Empty;
            bool CuscodeValido = false;

            bool datosExtraidos = false;
            int tabular = 0;

            // Bucle principal
            while (true)
            {
                try
                {
                    log.Escribir("Buscando casos");

                BuscarCaso:
                    caso = _casoRepo.ConsultarPendiente();
                    if (caso == null)
                    {
                        log.Escribir("No hay casos pendientes para procesar.");

                        log.Escribir("Consultando caso a procesar");
                        caso = _casoRepo.ConsultarPendiente();
                    }

                    if (caso == null)
                    {
                        log.Escribir("No hay casos pendientes para procesar. Finalizando ejecución.");
                        Console.WriteLine("No hay casos pendientes para procesar. Finalizando ejecución.");
                        log.finalizar();

                        string Kill = @"
                                try 
                                {
                                    RunWait, taskkill /f /im ""AC Administrador de Clientes.exe""
                                }
                                catch e
                                {
                                    ; Manejo de errores si es necesario
                                }
                            ";
                        try
                        {
                            ahk.ExecRaw(Kill.ToString());
                            Thread.Sleep(2000);
                            var resultado = ahk.GetVar("result");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error al ejecutar el script: {ex.Message}");
                            throw new AHKException($"Error: de AHK .", ex);
                        }

                        Environment.Exit(0); // Terminar el programa si no hay casos pendientes
                        break; // No hay más casos pendientes, salir del bucle
                    }

                    _casoRepo.IncrementarIntento(caso.Id_Lecturabilidad);

                    // Obtener el remitente a comprobar 
                    remitente = (caso.CorreoaNotificar ?? string.Empty).Trim();

                    // Normalizar: bajar a minúsculas y eliminar espacios laterales
                    remitenteNorm = remitente.ToLowerInvariant();

                    // Si el remitente está en la lista -> NO insertar
                    if (!string.IsNullOrEmpty(remitente) && blockedSenders.Contains(remitenteNorm))
                    {
                        log.Escribir($"Remitente bloqueado detectado para Id {caso.Id_Lecturabilidad}: {remitente} -> No se insertará en DetalleLecturabilidad. Atención manual requerida.");
                        Console.WriteLine($"Remitente bloqueado ({remitente}) — se omite InsertarDetalle para Id {caso.Id_Lecturabilidad}.");

                        // marcar el caso para revisión manual en lugar de dejarlo como EXITOSO.

                        _casoRepo.ActualizarEstado(
                            caso.Id_Lecturabilidad,
                            nuevoEstado: "RECHAZADO",
                            observaciones: $"Remitente bloqueado ({remitente}). Atención manual requerida.",
                            ExtraccionCompleta: caso.ExtraccionCompleta
                        );
                        log.Escribir($"Caso {caso.Id_Lecturabilidad} marcado como PENDIENTE_MANUAL por remitente bloqueado.");

                        goto BuscarCaso;
                        //  salimos del flujo de inserción.
                    }

                    if (caso.Cuscode == "0" && caso.Identificacion == "0")
                    {
                        _casoRepo.ActualizarEstado(
                                    caso.Id_Lecturabilidad,
                                    nuevoEstado: "RECHAZADO",
                                    observaciones: $"Caso Rechazado por datos incompletos",
                                    ExtraccionCompleta: caso.ExtraccionCompleta
                                );

                        goto BuscarCaso;
                    }

                    if (caso.CantidadDelineas == 0)
                    {
                        caso.CantidadDelineas = 1;

                        // Persistir CantidadDelineas 
                        try
                        {
                            _casoRepo.ActualizarDatos(caso.Id_Lecturabilidad, "CantidadDelineas", null, caso.CantidadDelineas);

                        }
                        catch (Exception ex)
                        {
                            log.Escribir($"Error al insertar en base de datos, nota RR para SolicitudId={caso.Id_Lecturabilidad}: {ex.Message}");
                            Console.WriteLine($"Error al insertar en base de datos, nota RR para SolicitudId={caso.Id_Lecturabilidad}: {ex.Message}");
                            throw;
                        }
                    }

                    if (!string.IsNullOrEmpty(caso.Cuscode) && caso.Cuscode != "0" && caso.Cuscode != ""
                        && !string.IsNullOrEmpty(caso.DirigidoA) && caso.DirigidoA != "0" && caso.DirigidoA != ""
                        && !string.IsNullOrEmpty(caso.Identificacion) && caso.Identificacion != "0" && caso.Identificacion != ""
                        && !string.IsNullOrEmpty(caso.CorreoaNotificar) && caso.CorreoaNotificar != "0" && caso.CorreoaNotificar != "")
                    {

                        string cus = caso.Cuscode;
                        BusquedaCuscode = ACHelper.IsValidCuscode(cus);

                        if (BusquedaCuscode)
                        {
                            log.Escribir("Caso ya cuenta con datos completos");
                            Console.WriteLine("Caso ya cuenta con datos completos");
                            caso.ExtraccionCompleta = true;

                            _casoRepo.ActualizarEstado(
                                caso.Id_Lecturabilidad,
                                nuevoEstado: "EXITOSO",
                                observaciones: $"Caso ya cuenta con datos completos",
                                ExtraccionCompleta: caso.ExtraccionCompleta
                            );
                            log.Escribir($"Caso: {caso.Id_Lecturabilidad} marcado como EXITOSO.");
                            Thread.Sleep(2000);

                            // remitente NO bloqueado -> proceder con la inserción
                            try
                            {
                                int filas = _actualizarCasoRepo.InsertarDetalle(caso.Id_Lecturabilidad);
                                Console.WriteLine($"InsertarDetalle: filas afectadas = {filas}");
                                log.Escribir($"InsertarDetalle: filas afectadas = {filas}");
                                goto BuscarCaso;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Fallo InsertarDetalle: " + ex.Message);
                                log.Escribir("Fallo InsertarDetalle: " + ex.Message);
                                throw;
                            }
                            
                        }
                    }
                    
                    

                    caso.ExtraccionCompleta = false;
                    bool procesoCompleto = false; // Inicializar como falso
                    int intentos = 0; // Contador de intentos

                    while (!procesoCompleto && intentos < _maxLoginAttempts)
                    {
                        try
                        {
                            log.Escribir($"Procesando caso: {caso.Id_Lecturabilidad} intento: {intentos}");

                            //iniciar autohotkey
                            string result = null;

                            //login RR 1
                            log.Escribir("Haciendo login en AC");

                            string loginAC = $@"
                                        global result
                                        RutaImagenGlobal := ""{RutaImagenGlobal}""

                                     try 
                                     {{

                                        RunWait, taskkill /f /im ""AC Administrador de Clientes.exe""

		                                Run, ""C:\Program Files (x86)\AC Administración de Clientes\AC Administrador de Clientes.exe""
		                                Sleep, 5000

		                                SetTitleMatchMode, 2	
		                                WinActivate, ""AC Administración de Clientes""
		                                WinGet, window_id, ID, AC Administración de Clientes
                                        WinActivate, ahk_id %window_id%
		                                Sleep, 3000
                                     
		                                If !FileExist( ""C:\Desarrollos-BI\ImagenesRobots\img_LogoClaroAC.PNG"" )
		                                {{
			                                ;MsgBox, NO EXISTE LA IMAGEN img_LogoAC.PNG
			                                throw Exception(""Image Doesn't Exist: img_LogoAC.PNG"")
		                                }}

		                                Loop, 120
		                                {{
			                                CoordMode, Pixel, Window
                                            rutaImg := RutaImagenGlobal . ""img_LogoClaroAC.PNG""
			                                ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %rutaImg%
			                                If ErrorLevel = 0
			                                {{
			
			                                    ; Establece el modo de coincidencia del título
				                                SetTitleMatchMode, 2  ; Coincidencia parcial del título

				                                ; Minimizar la ventana cuyo título contiene ""RPA_Validación_Datos_AC""
				                                WinMinimize, RPA_Validación_Datos_AC			   

				                                Click, %FoundX%, %FoundY%, Left, 2
				                                Sleep, 2000
				                                Break
			                                }}
			                                Sleep, 1200
		                                }}

		                                If ErrorLevel != 0
		                                {{
			                                throw Exception(""Image Not Found: img_LogoClaroAC.PNG"")
		                                }}       	    

                                        ;se inicia el logueo en el ac 
		                                MouseMove, 190, 272
		                                Sleep, 300
		                                MouseClick, Left, 190, 272, 1
		                                SendRaw, {clave}
		                                Sleep, 600
		                                Send, {{Tab}}
		                                Sleep, 600
		                                Send, {{Down}}
		                                Sleep, 600
		                                Send, {{Enter}}
		                                Sleep, 2000
		                                Send, {{Enter}}
                                        Sleep, 4000

                                        Img := RutaImagenGlobal . ""ERROR_CREDENCIALES.PNG""
                                        ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %Img%
			                            If ErrorLevel = 0
			                            {{
                                            result := ""ERROR_CREDENCIALES""
                                            return
			                            }}

		                                Sleep, 8000

		                                result := ""OK""
	                                }}
	                                catch e
	                                {{
		                                result := ""(AHK "" e.What "": "" e.Line "") "" e.Message
	                                }}

	                                return
                                        
                                                            ";

                            try
                            {
                                ahk.ExecRaw(loginAC.ToString());
                                Thread.Sleep(2000);
                                result = ahk.GetVar("result");
                                if (result == "OK") // Asegúrate de que el resultado sea una cadena "true"
                                {
                                    Console.WriteLine("OK");
                                    log.Escribir("login OK");
                                }
                                else if (result == "ERROR_CREDENCIALES")
                                {
                                    log.Escribir("Error credenciales incorrectas");
                                    _casoRepo.IncrementarIntento(caso.Id_Lecturabilidad);

                                    _emailSender.SendMailException("Error credenciales incorrectas", caso?.Id_Lecturabilidad, "Error credenciales incorrectas");
                                    log.finalizar();

                                    // Terminar el programa ya que es un error critico
                                    Environment.Exit(1);

                                }
                                else if (result.Contains("Image Doesn't Exist: img_LogoAC.PNG"))
                                {
                                    log.Escribir("Error critico! imagen: img_LogoAC.PNG no existe en la ruta.");
                                    log.Escribir($"Verifique que la imagen se encuentra en la ruta: C:\\Desarrollos-BI\\ImagenesRobots\\");
                                    Console.WriteLine("Error critico! imagen: img_LogoAC.PNG no existe en la ruta.");
                                    Console.WriteLine($"Verifique que la imagen se encuentra en la ruta: C:\\Desarrollos-BI\\ImagenesRobots\\");
                                    _emailSender.SendMailException("Error critico! imagen: img_LogoAC.PNG no existe en la ruta", caso?.Id_Lecturabilidad, "Verifique que la imagen se encuentra en la ruta: C:\\Desarrollos-BI\\ImagenesRobots\\");
                                    log.finalizar();

                                    // Terminar el programa ya que es un error critico
                                    Environment.Exit(1);
                                    //throw new ACImgException($"Error critico: imagen: img_LogoAC.PNG no existe en la ruta.");
                                }
                                else if (result.Contains("Image Not Found: img_LogoClaroAC.PNG"))
                                {
                                    log.Escribir("Error! imagen: img_LogoAC.PNG no encontrada.");
                                    Console.WriteLine("Error! imagen: img_LogoAC.PNG no encontrada.");
                                    throw new ACImgException("No se encontro imagen: img_LogoClaroAC.PNG");
                                }
                                else if (result == "")
                                {
                                    log.Escribir("Error critico! Al ejecutar script ahk");
                                    throw new AHKException($"Error: fallo al ejecutar AHK.");
                                }
                                else
                                {
                                    throw new ACException($"Error: De login AC.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error al ejecutar el script: {ex.Message}");
                                throw new ACException($"Error: De login AC.", ex);
                            }

                            string cus = caso.Cuscode;
                            if (!string.IsNullOrEmpty(caso.Cuscode) && caso.Cuscode != "" && caso.Cuscode != "0")
                            {
                                BusquedaCuscode = ACHelper.IsValidCuscode(cus);
                            } 

                            // consulta de suscriptor 
                            if (BusquedaCuscode)
                            {
                                // consulta de suscriptor por Custcode
                                Thread.Sleep(3000);
                                string Consulta = $@"
                                            global result
                                            lineaEncontrada := 0  ; 0 = No encontrada, 1 = Encontrada
                                            validacion := 0

                                            try
	                                            {{
		                                            WinActivate, ""AC Administración de Clientes""
		                                            Sleep, 1200

		                                            WinMaximize, ""AC Administración de Clientes""
		                                            Sleep, 1200

		                                            If !FileExist(""C:\Desarrollos-BI\ImagenesRobots\img_CUST_CODE.PNG"")
		                                            {{
			                                            throw Exception(""Image Doesn't Exist: img_CUST_CODE.PNG"")
		                                            }}

		                                            Loop, 120
		                                            {{
			                                            WinActivate, ""AC Administración de Clientes""
			                                            Sleep, 500

			                                            CoordMode, Pixel, Window
			                                            ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_CUST_CODE.PNG
			                                            If ErrorLevel = 0
			                                            {{
				                                            ;MsgBox, encuentra Imagen img_CUST_CODE
				                                            Click, %FoundX%, %FoundY% Left, 1
				                                            Sleep, 2000
				                                            Break
			                                            }}
			                                            Sleep, 500
		                                            }}

		                                            If ErrorLevel != 0
		                                            {{
			                                            ;MsgBox, no funciona la imagen img_CUST_CODE
			                                            throw Exception(""Image Not Found: img_CUST_CODE.PNG"")
		                                            }}

		                                            ;~ Damos click en campo CRITERIOS para pasarle el valor de su variable
		                                            MouseMove, 20, 255
		                                            ;MouseMove, 30, 227
		                                            Sleep, 300
		                                            MouseClick, Left, 73, 255, 1
		                                            ;MouseClick, Left, 30, 227, 1
		                                            Sleep, 300
		                                            SendRaw, {caso.Cuscode}
		                                            Sleep, 300

		                                            If !FileExist(""C:\Desarrollos-BI\ImagenesRobots\img_BuscarCUN.PNG"")
		                                            {{
			                                            ;MsgBox, No existe Imagen img_BuscarCUN
			                                            throw Exception(""Image Doesn't Exist: img_BuscarCUN.PNG"")
		                                            }}

		                                            Loop, 120
		                                            {{
			                                            WinActivate, ""AC Administración de Clientes""
			                                            Sleep, 500

			                                            CoordMode, Pixel, Window
			                                            ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_BuscarCUN.PNG
			                                            If ErrorLevel = 0
			                                            {{
				                                            ;MsgBox, encuentra Imagen img_BuscarCUN
				                                            Click, %FoundX%, %FoundY% Left, 1
				                                            Sleep, 2000
				                                            Break
			                                            }}
			                                            Sleep, 2000
		                                            }}

		                                            If ErrorLevel != 0
		                                            {{
			                                            ;MsgBox, no funciona la imagen img_BuscarCUN
			                                            throw Exception(""Image Not Found: img_BuscarCUN"")
		                                            }}

                                                    ;comprobar cuenta
                                                    Img := RutaImagenGlobal . ""No_Existe_Informacion.PNG""
                                                    ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %Img%
                                                    if (ErrorLevel = 0) {{
                                                        ; imagen encontrada -> devolver estado y no continuar
                                                        result := ""No_Existe_Informacion""
                                                        ;MsgBox, No_Existe_Informacion
                                                        return
                                                    }} 

		                                            Sleep, 3000
		                                            ;~ Damos click en ""Cuenta"" Primera fila
		                                            MouseMove, 390, 135
		                                            Sleep, 600
		                                            MouseClick, Left, 390, 135, 1
		                                            Sleep, 5000

                                                    ; Comprobar si la linea es Prepago
                                                    Imagen := RutaImagenGlobal . ""Prepago.PNG""
                                                    ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %Imagen%
                                                    if (ErrorLevel = 0)
                                                    {{
                                                        ; Imagen encontrada: establecer la bandera y CONTINUAR el script
                                                        lineaEncontrada := 1
    
                                                        ;MsgBox, Línea de Prepago detectada. Continuando...
                                                    }}
                                                    Sleep, 3000

                                                    if (lineaEncontrada = 0)
                                                    {{
                                                        ; Comprobar si la linea es Postpago
                                                        Imagen := RutaImagenGlobal . ""Postpago.PNG""
                                                        ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %Imagen%
                                                        if (ErrorLevel = 0)
                                                        {{
                                                            ; Imagen encontrada: establecer la bandera y CONTINUAR el script
                                                            lineaEncontrada := 1
                                                            ;MsgBox, Línea de Postpago detectada. Continuando...
                                                        }}
                                                    }}

                                                    Sleep, 3000

                                                    if (lineaEncontrada = 0)
                                                    {{
                                                        ; Si la bandera sigue en 0, ninguna de las dos imágenes fue encontrada.
                                                        result := ""FALLO: Tipo de línea no detectado (ni Prepago ni Postpago).""
                                                        ;MsgBox, Fallo: No se detectó ni Prepago ni Postpago. Saliendo del script.
                                                        return ; Detiene la ejecución del script aquí y retorna
                                                    }}

                                                    ;Comprobar si la linea esta desactivada
                                                    Img2 := RutaImagenGlobal . ""Linea_Desactivada.PNG""
                                                    ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %Img2%
                                                    if (ErrorLevel = 0) {{
                                                        ; imagen encontrada -> devolver estado y no continuar
                                                        result := ""Linea_Desactivada""
                                                        ;MsgBox, Linea_Desactivada
                                                        return
                                                    }} 
                                                    Sleep, 3000

		                                            ;~ Damos click en ""Cuenta"" Primera fila
		                                            MouseMove, 390, 305
		                                            Sleep, 600
		                                            MouseClick, Left, 390, 305, 1
		                                            Sleep, 3000

		                                            If !FileExist(""C:\Desarrollos-BI\ImagenesRobots\img_BotonConsultarCUN.PNG"")
		                                            {{
			                                            throw Exception(""Image Doesn't Exist: img_BotonConsultarCUN.PNG"")
		                                            }}

		                                            Loop, 120
		                                            {{
			                                            WinActivate, ""AC Administración de Clientes""
			                                            Sleep, 500

			                                            CoordMode, Pixel, Window
			                                            ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_BotonConsultarCUN.PNG
			                                            If ErrorLevel = 0
			                                            {{
				                                            ;MsgBox, encuentra Imagen img_BotonConsultarCUN
				                                            Click, %FoundX%, %FoundY% Left, 1
				                                            Sleep, 4000
				                                            Break
			                                            }}
			                                            Sleep, 2000
		                                            }}

		                                            If ErrorLevel != 0
		                                            {{
			                                            throw Exception(""Image Not Found: img_BotonConsultarCUN"")
		                                            }}

		                                            Sleep, 5000

		                                            ;If !FileExist(""C:\Desarrollos-BI\ImagenesRobots\img_AceptarLineaProceso.PNG"")
		                                            ;{{
			                                            ;throw Exception(""Image Doesn't Exist: img_AceptarLineaProceso.PNG"")
		                                            ;}}
		

		                                            ;Loop, 120
		                                            ;{{
			                                           ; WinActivate, ""AC Administración de Clientes""
			                                            ;Sleep, 500

			                                            ;CoordMode, Pixel, Window
			                                            ;ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_AceptarLineaProceso.PNG
			                                            ;If ErrorLevel = 0
			                                            ;{{
				                                           ; Click, %FoundX%, %FoundY% Left, 1
				                                           ; Sleep, 2000
				                                           ; Break
			                                            ;}}
			                                            ;else
			                                            ;{{
				                                           ; Break
			                                            ;}}
			                                            ;Sleep, 1000
		                                            ;}}

                                                    Sleep, 2000

                                                    ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\ValidacionBarrio.PNG
			                                        If ErrorLevel = 0
			                                        {{
				                                            ;Click, %FoundX%, %FoundY% Left, 1
                                                            validacion := 1
                                                            Sleep, 2000
                                                            Send, {{Enter}}
				                                            Sleep, 3000  
				                                            
			                                        }}
			                                         
			                                        Sleep, 2000

		                                            If !FileExist(""C:\Desarrollos-BI\ImagenesRobots\img_BotonAceptarCampoBarrio.PNG"")
		                                            {{
			                                            throw Exception(""Image Doesn't Exist: img_BotonAceptarCampoBarrio.PNG"")
		                                            }}

                                                    if (validacion = 1)
                                                    {{
                                                        ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_BotonAceptarCampoBarrio.PNG
			                                            If ErrorLevel = 0
			                                            {{
				                                                ;Click, %FoundX%, %FoundY% Left, 1
                                                                Send, {{Enter}}
				                                                Sleep, 2000
                                                                Send, {{Tab 36}}
                                                                Sleep, 3000
				                                            
			                                            }}
			                                         
			                                            Sleep, 1000
                                                    }}
                                                    else 
                                                    {{
                                                        ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_BotonAceptarCampoBarrio.PNG
			                                            If ErrorLevel = 0
			                                            {{
				                                                ;Click, %FoundX%, %FoundY% Left, 1
                                                                Send, {{Enter}}
				                                                Sleep, 2000
                                                                Send, {{Tab 37}}
                                                                Sleep, 3000
				                                            
			                                            }}
			                                         
			                                            Sleep, 1000
                                                    }}     	                                            

                                                    ;quitar mensajes informativos 
		                                              ;por si sale mensaje de confirmacion 
		                                            ;Send, {{Enter}}
                                                    ;Sleep, 3000
                                                    ;Send, {{Enter}}
                                                    ;Sleep, 3000
                                                    ;Send, {{Enter}}		
		                                            ;Sleep, 3000

		                                            ;If !FileExist(""C:\Desarrollos-BI\ImagenesRobots\img_CiudadDptoIncorrectos.PNG"")
		                                            ;{{
			                                            ;throw Exception(""Image Doesn't Exist: img_CiudadDptoIncorrectos.PNG"")
		                                            ;}}

		                                            ;Loop, 30
		                                            ;{{
			                                            ;WinActivate, ""Validación Ciudad/Departamento""
			                                           ; Sleep, 500

			                                            ;CoordMode, Pixel, Window
			                                            ;ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_CiudadDptoIncorrectos.PNG
			                                            ;If ErrorLevel = 0
			                                            ;{{
				                                            ;Click, %FoundX%, %FoundY% Left, 1
				                                            ;Sleep, 2000
				                                            ;Break
			                                            ;}}
			                                            ;else
			                                            ;{{
				                                            ;Break
			                                            ;}}
			                                            ;Sleep, 1000
		                                            ;}}

		                                            ;If !FileExist(""C:\Desarrollos-BI\ImagenesRobots\img_BotonAceptarFechaCum.PNG"")
		                                            ;{{
			                                            ;throw Exception(""Image Doesn't Exist: img_BotonAceptarFechaCum.PNG"")
		                                            ;}}

		                                            ;Loop, 30
		                                            ;{{
			                                            ;WinActivate, ""AC Administración de Clientes""
			                                            ;Sleep, 500

			                                            ;CoordMode, Pixel, Window
			                                            ;ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_BotonAceptarFechaCum.PNG
			                                            ;If ErrorLevel = 0
			                                            ;{{
				                                            ;Click, %FoundX%, %FoundY% Left, 1
				                                            ;Sleep, 2000
				                                            ;Break
			                                            ;}}
			                                            else
			                                            ;{{
				                                            ;Break
			                                            ;}}
			                                            ;Sleep, 1000
		                                            ;}}


		                                            ;If !FileExist(""C:\Desarrollos-BI\ImagenesRobots\img_FalloConexionBD.PNG"")
		                                            ;{{
			                                            ;throw Exception(""Image Doesn't Exist: img_FalloConexionBD.PNG"")
		                                            ;}}

		                                            ;Loop, 30
		                                            ;{{
			                                            ;WinActivate, ""AC Administración de Clientes""
			                                            ;Sleep, 500

			                                            ;CoordMode, Pixel, Window
			                                            ;ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_FalloConexionBD.PNG
			                                            ;If ErrorLevel = 0
			                                            ;{{
				                                            ;MsgBox, encuentra Imagen img_FalloConexionBD
				                                            ;Send, {{Enter}}
				                                            ;Sleep, 2000
				                                            ;Break
			                                            ;}}
			                                            ;else
			                                            ;{{
				                                            ;Break
			                                            ;}}
			                                            ;Sleep, 1000
		                                            ;}}

		                                            result := ""OK""
		                                            ;MsgBox, Valor variable Resultado %Resultado%

	                                            }}
	                                            catch e
	                                            {{
		                                            result := ""(AHK "" e.What "": "" e.Line "") "" e.Message
	                                            }}

	                                            return
                                                            
                                ";


                                try
                                {
                                    ahk.ExecRaw(Consulta.ToString());
                                    Thread.Sleep(2000);
                                    result = ahk.GetVar("result");
                                    if (result == "OK") // Asegúrate de que el resultado sea una cadena "true"
                                    {
                                        Console.WriteLine("OK");
                                    }
                                    else if (result.Contains("Image Doesn't Exist"))
                                    {
                                        log.Escribir("Error critico! imagen no existe en la ruta.");
                                        log.Escribir($"Verifique que las imagenes se encuentran en la ruta: C:\\Desarrollos-BI\\ImagenesRobots\\");
                                        Console.WriteLine("Error critico! imagen no existe en la ruta.");
                                        Console.WriteLine($"Verifique que las imagenes se encuentran en la ruta: C:\\Desarrollos-BI\\ImagenesRobots\\");
                                        _emailSender.SendMailException("Error critico! imagen no encontrada", caso?.Id_Lecturabilidad, "Verifique que la imagenes se encuentran en la ruta: C:\\Desarrollos-BI\\ImagenesRobots\\");
                                        log.finalizar();

                                        // Terminar el programa ya que es un error critico
                                        Environment.Exit(1);
                                        //throw new ACImgException($"Error critico: imagen: img_LogoAC.PNG no existe en la ruta.");
                                    }
                                    else if (result.Contains("Image Not Found"))
                                    {
                                        log.Escribir("Error! imagen no encontrada.");
                                        Console.WriteLine("Error! imagen no encontrada.");
                                        throw new ACImgException("Error! imagen no encontrada");
                                    }
                                    else if (result == "")
                                    {
                                        log.Escribir("Error critico! Al ejecutar script ahk");
                                        throw new AHKException($"Error: fallo al ejecutar AHK.");
                                    }
                                    else if (result.Contains("No_Existe_Informacion"))
                                    {
                                        log.Escribir("No existe información para el Cuscode proporcionado.");
                                        Console.WriteLine("No existe información para el Cuscode proporcionado.");

                                        /*_casoRepo.ActualizarEstado(
                                            caso.Id_Lecturabilidad,
                                            nuevoEstado: "RECHAZADO",
                                            observaciones: $"No existe informacion para la identificacion consultada",
                                            ExtraccionCompleta: caso.ExtraccionCompleta
                                        );*/
                                        BusquedaCuscode = false;
                                        goto BuscarCedula;
                                    }
                                    else if(result.Contains("Linea_Desactivada"))
                                    {
                                        log.Escribir("La línea asociada al Cuscode está desactivada.");
                                        Console.WriteLine("La línea asociada al Cuscode está desactivada.");
                                        _casoRepo.ActualizarEstado(
                                            caso.Id_Lecturabilidad,
                                            nuevoEstado: "RECHAZADO",
                                            observaciones: $"La linea asociada a la Cuscode consultada esta desactivada",
                                            ExtraccionCompleta: caso.ExtraccionCompleta
                                        );
                                        goto BuscarCaso;
                                    }
                                    else if (result.Contains("No_Tiene_Linea"))
                                    {
                                        log.Escribir("El Cuscode consultado no tiene línea asociada.");
                                        Console.WriteLine("El Cuscode consultado no tiene línea asociada.");
                                        _casoRepo.ActualizarEstado(
                                            caso.Id_Lecturabilidad,
                                            nuevoEstado: "RECHAZADO",
                                            observaciones: $"El Cuscode consultado no tiene linea asociada",
                                            ExtraccionCompleta: caso.ExtraccionCompleta
                                        );
                                        goto BuscarCaso;
                                    }
                                    else if (result.Contains("FALLO: Tipo de línea no detectado (ni Prepago ni Postpago)."))
                                    {
                                        log.Escribir("La línea asociada al Cuscode no tiene tipo de línea (Prepago/Postpago).");
                                        Console.WriteLine("La línea asociada al Cuscode no tiene tipo de línea (Prepago/Postpago).");
                                        _casoRepo.ActualizarEstado(
                                            caso.Id_Lecturabilidad,
                                            nuevoEstado: "RECHAZADO",
                                            observaciones: $"La linea sin datos",
                                            ExtraccionCompleta: caso.ExtraccionCompleta
                                        );
                                        goto BuscarCaso;

                                    }
                                    else
                                    {
                                        throw new ACException($"Error: En AC.");
                                    }

                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error al ejecutar el script: {ex.Message}");
                                    throw new ACException($"Error: En AC.", ex);
                                }

                                Thread.Sleep(2000);
                                string ventana = @"
                                            SetTitleMatchMode, 2

                                            WinTitle := ""AC Administración de Clientes""

                                            ; Intentar activar y traer al frente la ventana
                                            WinActivate, %WinTitle%
                                            WinWaitActive, %WinTitle%, , 5
                                            ; Forzar Foreground
                                            WinExistId := WinExist(WinTitle)
                                            if (WinExistId)
                                            {
                                                ; asegurar foreground (DLL)
                                                DllCall(""SetForegroundWindow"", ""ptr"", WinExistId)
                                                ; temporalmente arriba
                                                WinSet, AlwaysOnTop, On, ahk_id %WinExistId%
                                            }
                                            Sleep, 300
                                            ";

                                try
                                {
                                    ahk.ExecRaw(ventana.ToString());
                                    Thread.Sleep(2000);
                                    result = ahk.GetVar("result");
                                    if (result == "OK") // Asegúrate de que el resultado sea una cadena "true"
                                    {
                                        Console.WriteLine("OK");
                                    }
                                    else
                                    {
                                        throw new ACException($"Error: En AC.");
                                    }

                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error al ejecutar el script: {ex.Message}");
                                    throw new ACException($"Error: En AC.", ex);
                                }

                                if (string.IsNullOrEmpty(caso.DirigidoA) || caso.DirigidoA == "0")
                                {
                                    // Captura de nombre
                                    Thread.Sleep(3000);
                                    string Nombre = @"
                                            global result
                                            global nombre
                                            global apellido
                                            nombre := """"
                                            apellido := """"
                                            SetTitleMatchMode, 2

                                            WinActivate, ""AC Administración de Clientes""
		                                    Sleep, 1200

                                            ;usar coordenadas relativas a la ventana
                                            CoordMode, Mouse, Window
                                            CoordMode, Pixel, Window

                                            ; posicionar sobre la celda de la Nombre y copiar
                                            ;MouseMove, 278, 164
                                            Sleep, 500
                                            ;Click, left, 278, 164
                                            Send, {Tab}
                                            Sleep, 1000
                                            Send, +{Right 40}
                                            Sleep, 2000
                                            Clipboard := """"               ; limpiar clipboard
                                            SendInput, ^c
                                            ClipWait, 2                 ; espera hasta 2s (ajusta)
                                            if ErrorLevel
                                            {{
                                                ; no copió: retry pequeño y volver a intentar una copia directa
                                                Sleep, 150
                                                SendInput, ^c
                                                ClipWait, 1
                                            }}

                                            current := RegExReplace(Clipboard, ""^\s+|\s+$"", """") ; trim
                                            ;MsgBox, copiado  %current%
                                            nombre := current

                                            ; posicionar sobre la celda de la Apellido y copiar
                                            ;MouseMove, 278, 189
                                            Sleep, 500
                                            ;Click, left, 278, 189
                                            Send, {Tab}
                                            Sleep, 1000
                                            Send, +{Right 40}
                                            Sleep, 2000
                                            Clipboard := """"               ; limpiar clipboard
                                            SendInput, ^c
                                            ClipWait, 2                 ; espera hasta 2s (ajusta)
                                            if ErrorLevel
                                            {{
                                                ; no copió: retry pequeño y volver a intentar una copia directa
                                                Sleep, 150
                                                SendInput, ^c
                                                ClipWait, 1
                                            }}

                                            current := RegExReplace(Clipboard, ""^\s+|\s+$"", """") ; trim
                                            ;MsgBox, copiado %current%
                                            apellido := current

                                            ; Normalizar variables (evitar pipes en los valores)
                                            if (nombre = """")
                                                nombre := """"
                                            else
                                                nombre := StrReplace(nombre, ""|"", ""-"")  ; por seguridad, quitar pipes

                                            if (apellido = """")
                                                apellido := """"
                                            else
                                                apellido := StrReplace(apellido, ""|"", ""-"")

                                            if (nombre != """" && apellido != """")
                                            {{
                                                result := ""OK|"" . nombre . ""|"" . apellido
                                                return
                                            }}
                                            else
                                            {{
                                                result := ""FAIL_NOTFOUND""
                                                return
                                            }}
                                        
                                        ";

                                    try
                                    {
                                        ahk.ExecRaw(Nombre.ToString());
                                        Thread.Sleep(2000);
                                        result = ahk.GetVar("result");
                                        tabular = 1;
                                        if (result.Contains("OK")) // Asegúrate de que el resultado sea una cadena "true"
                                        {
                                            Console.WriteLine("OK");
                                            log.Escribir("Nombre extraido...");

                                            // Separar por '|'
                                            var parts = result.Split('|');
                                            status = parts.Length > 0 ? parts[0] : "";
                                            nombre = parts.Length > 1 ? parts[1] : "";
                                            apellido = parts.Length > 2 ? parts[2] : "";
                                            nombreCompleto = nombre + " " + apellido;
                                            // Guardar en DTO/objeto
                                            caso.DirigidoA = nombreCompleto;

                                            // Persistir identificacion 
                                            try
                                            {
                                                if (!string.IsNullOrEmpty(caso.DirigidoA) && caso.DirigidoA != "0")
                                                {
                                                    _casoRepo.ActualizarDatos(caso.Id_Lecturabilidad, "DirigidoA", caso.DirigidoA);

                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                log.Escribir($"Error al insertar en base de datos, nota RR para SolicitudId={caso.Id_Lecturabilidad}: {ex.Message}");
                                                Console.WriteLine($"Error al insertar en base de datos, nota RR para SolicitudId={caso.Id_Lecturabilidad}: {ex.Message}");
                                                throw;
                                            }
                                        }
                                        else if (result.Contains("FAIL_NOTFOUND"))
                                        {
                                            log.Escribir("No se pudo extraer el nombre");
                                            Console.WriteLine("No se pudo extraer el nombre");
                                            throw new ACException($"Error: En AC.");

                                        }
                                        else if (result == "")
                                        {
                                            log.Escribir("Error critico! Al ejecutar script ahk");
                                            throw new AHKException($"Error: fallo al ejecutar AHK.");
                                        }
                                        else
                                        {
                                            throw new ACException($"Error: En AC.");
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Error al ejecutar el script: {ex.Message}");
                                        throw new ACException($"Error: En AC.");
                                    }

                                }

                                if (string.IsNullOrEmpty(caso.Identificacion) || caso.Identificacion == "0")
                                {

                                    // Captura de identicacion
                                    Thread.Sleep(3000);
                                    string scriptIdentificacion = $@"
                                            global result
                                            global identificacion
                                            global tabular
                                            tabular := ""{tabular}""
                                            identificacion := """"
                                            SetTitleMatchMode, 2

                                            WinActivate, ""AC Administración de Clientes""
		                                    Sleep, 1200

                                            ;usar coordenadas relativas a la ventana
                                            CoordMode, Mouse, Window
                                            CoordMode, Pixel, Window

                                            ; posicionar sobre la celda de la Nombre y copiar
                                            ;MouseMove, 278, 213
                                            Sleep, 500
                                            Clipboard := """"               ; limpiar clipboard
                                            Sleep, 1000
                                            ;MouseClick, left, 278, 213
                                            if (tabular = 0)
                                            {{
                                                Send, {{Tab 3}}
                                            }}
                                            else 
                                            {{
                                                Send, {{Tab}}
                                            }}
                                            
                                            Sleep, 1000
                                            Send, +{{Right 20}}
                                            Sleep, 1000
                                            MouseClick, right, 278, 213
                                            Sleep, 2000
                                            Send, {{c}}
                                            Sleep, 500
                                            
                                            ClipWait, 2                 ; espera hasta 2s (ajusta)
                                            if ErrorLevel
                                            {{
                                                ; no copió: retry pequeño y volver a intentar una copia directa
                                                Sleep, 150
                                                result := ""FAIL_CLIPBOARD""
                                                return
                                            }}

                                            current := RegExReplace(Clipboard, ""^\s+|\s+$"", """") ; trim
                                            ;MsgBox, copiado %current%
                                            identificacion := current

                                            ; Normalizar variables (evitar pipes en los valores)
                                            if (identificacion = """")
                                                identificacion := """"
                                            else
                                                identificacion := StrReplace(identificacion, ""|"", ""-"")  ; por seguridad, quitar pipes

                                            if (identificacion != """")
                                            {{
                                                result := ""OK|"" . identificacion
                                                return
                                            }}
                                            else
                                            {{
                                                result := ""FAIL_NOTFOUND""
                                                return
                                            }}
                                        
                                        ";

                                    try
                                    {
                                        ahk.ExecRaw(scriptIdentificacion.ToString());
                                        Thread.Sleep(2000);
                                        result = ahk.GetVar("result");
                                        tabular = 2;
                                        if (result.Contains("OK")) // Asegúrate de que el resultado sea una cadena "true"
                                        {
                                            Console.WriteLine("OK");
                                            log.Escribir("identificacion extraida...");

                                            // Separar por '|'
                                            var parts = result.Split('|');
                                            status = parts.Length > 0 ? parts[0] : "";
                                            identificacion = parts.Length > 1 ? parts[1] : "";

                                            // Guardar en DTO/objeto
                                            caso.Identificacion = identificacion;

                                            // Persistir identificacion 
                                            try
                                            {
                                                if (!string.IsNullOrEmpty(caso.Identificacion) && caso.Identificacion != "0")
                                                {
                                                    _casoRepo.ActualizarDatos(caso.Id_Lecturabilidad, "Identificacion", caso.Identificacion);

                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                log.Escribir($"Error al insertar en base de datos, nota RR para SolicitudId={caso.Id_Lecturabilidad}: {ex.Message}");
                                                Console.WriteLine($"Error al insertar en base de datos, nota RR para SolicitudId={caso.Id_Lecturabilidad}: {ex.Message}");
                                                throw;
                                            }
                                        }
                                        else if (result.Contains("FAIL_NOTFOUND"))
                                        {
                                            log.Escribir("No se pudo extraer la identificacion");
                                            Console.WriteLine("No se pudo extraer la identificacion");
                                            throw new ACException($"Error: En AC.");
                                        }
                                        else if (result == "")
                                        {
                                            log.Escribir("Error critico! Al ejecutar script ahk");
                                            throw new AHKException($"Error: fallo al ejecutar AHK.");
                                        }
                                        else
                                        {
                                            throw new ACException($"Error: En AC.");
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Error al ejecutar el script: {ex.Message}");
                                        throw new ACException($"Error: En AC.");
                                    }
                                }

                                if (string.IsNullOrEmpty(caso.CorreoaNotificar) || caso.CorreoaNotificar == "0")
                                {
                                    // Captura de correo
                                    Thread.Sleep(3000);
                                    string scriptcorreo = $@"
                                            global result
                                            global correo
                                            global tabular 
                                            tabular := ""{tabular}""
                                            correo := """"
                                            SetTitleMatchMode, 2

                                            WinActivate, ""AC Administración de Clientes""
		                                    Sleep, 1200

                                            ;usar coordenadas relativas a la ventana
                                            CoordMode, Mouse, Window
                                            CoordMode, Pixel, Window

                                            ; posicionar sobre la celda de la Nombre y copiar
                                            ;MouseMove, 277, 322
                                            Sleep, 500
                                            ;MouseClick, left, 277, 322
                                            if (tabular = 0)
                                            {{
                                                Send, {{Tab 8}}
                                            }}
                                            else if (tabular = 1)
                                            {{
                                                Send, {{Tab 6}}
                                            }}
                                            else if (tabular = 2)
                                            {{
                                                Send, {{Tab 5}}
                                            }}
                                            Sleep, 1000
                                            Send, +{{Right 20}}
                                            Sleep, 2000
                                            Clipboard := """"               ; limpiar clipboard
                                            Send, ^c
                                            ClipWait, 2                 ; espera hasta 2s (ajusta)
                                            if ErrorLevel
                                            {{
                                                ; no copió: retry pequeño y volver a intentar una copia directa
                                                Sleep, 150
                                                Send, ^c
                                                ClipWait, 1
                                            }}

                                            current := RegExReplace(Clipboard, ""^\s+|\s+$"", """") ; trim
                                            ;MsgBox, copiado pqr %current%
                                            correo := current

                                            ; Normalizar variables (evitar pipes en los valores)
                                            if (correo = """")
                                                correo := """"
                                            else
                                                correo := StrReplace(correo, ""|"", ""-"")  ; por seguridad, quitar pipes

                                            if (correo != """")
                                            {{
                                                result := ""OK|"" . correo
                                                return
                                            }}
                                            else
                                            {{
                                                result := ""FAIL_NOTFOUND""
                                                return
                                            }}
                                        
                                        ";

                                    try
                                    {
                                        ahk.ExecRaw(scriptcorreo.ToString());
                                        Thread.Sleep(2000);
                                        result = ahk.GetVar("result");
                                        if (result.Contains("OK")) // Asegúrate de que el resultado sea una cadena "true"
                                        {
                                            Console.WriteLine("OK");
                                            log.Escribir("PQR encontrada...");

                                            // Separar por '|'
                                            var parts = result.Split('|');
                                            status = parts.Length > 0 ? parts[0] : "";
                                            correo = parts.Length > 1 ? parts[1] : "";

                                            // Guardar en DTO/objeto
                                            caso.CorreoaNotificar = correo;

                                            // Persistir la correo 
                                            try
                                            {
                                                if (!string.IsNullOrEmpty(caso.CorreoaNotificar) && caso.CorreoaNotificar != "0")
                                                {
                                                    _casoRepo.ActualizarDatos(caso.Id_Lecturabilidad, "CorreoaNotificar", caso.CorreoaNotificar);
                                                    //goto BuscarNotasHistorico;

                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                log.Escribir($"Error al insertar en base de datos, nota RR para SolicitudId={caso.Id_Lecturabilidad}: {ex.Message}");
                                                Console.WriteLine($"Error al insertar en base de datos, nota RR para SolicitudId={caso.Id_Lecturabilidad}: {ex.Message}");
                                                throw;
                                            }
                                        }
                                        else if (result.Contains("FAIL_NOTFOUND"))
                                        {
                                            log.Escribir("No se pudo extraer la identificacion");
                                            Console.WriteLine("No se pudo extraer la identificacion");
                                            throw new ACException($"Error: En AC.");

                                        }
                                        else if (result == "")
                                        {
                                            log.Escribir("Error critico! Al ejecutar script ahk");
                                            throw new AHKException($"Error: fallo al ejecutar AHK.");
                                        }
                                        else
                                        {
                                            throw new ACException($"Error: En AC.");
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Error al ejecutar el script: {ex.Message}");
                                        throw new ACException($"Error: En AC.");
                                    }
                                }

                                if (caso.DirigidoA != "0" && caso.Identificacion != "0" && caso.CorreoaNotificar != "0")
                                {
                                    log.Escribir("Datos extraidos");
                                    Console.WriteLine("Datos extraidos");
                                    caso.ExtraccionCompleta = true;
                                    goto TerminarIntento;
                                }
                                else
                                {
                                    log.Escribir("Datos no pudieron ser extraidos");
                                    Console.WriteLine("Datos no pudieron ser extraidos");
                                    caso.ExtraccionCompleta = false;
                                    goto TerminarIntento;
                                }

                            }

                        BuscarCedula:
                            if (!BusquedaCuscode && caso.Identificacion != "0" && !string.IsNullOrEmpty(caso.Identificacion))
                            {

                                // consulta de suscriptor por identificacion
                                Thread.Sleep(3000);
                                string Consulta2 = $@"
                                            global result
                                            lineaEncontrada := 0
                                            validacion := 0

                                            try
	                                            {{
		                                            WinActivate, ""AC Administración de Clientes""
		                                            Sleep, 1200

		                                            WinMaximize, ""AC Administración de Clientes""
		                                            Sleep, 1200

		                                            If !FileExist(""C:\Desarrollos-BI\ImagenesRobots\cedula.PNG"")
		                                            {{
			                                            throw Exception(""Image Doesn't Exist: cedula.PNG"")
		                                            }}

		                                            Loop, 120
		                                            {{
			                                            WinActivate, ""AC Administración de Clientes""
			                                            Sleep, 500

			                                            CoordMode, Pixel, Window
			                                            ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\cedula.PNG
			                                            If ErrorLevel = 0
			                                            {{
				                                            ;MsgBox, encuentra Imagen cedula
				                                            Click, %FoundX%, %FoundY% Left, 1
				                                            Sleep, 2000
				                                            Break
			                                            }}
			                                            Sleep, 500
		                                            }}

		                                            If ErrorLevel != 0
		                                            {{
			                                            ;MsgBox, no funciona la imagen cedula
			                                            throw Exception(""Image Not Found: cedula.PNG"")
		                                            }}

		                                            ;~ Damos click en campo CRITERIOS para pasarle el valor de su variable
		                                            MouseMove, 20, 255
		                                            ;MouseMove, 30, 227
		                                            Sleep, 300
		                                            MouseClick, Left, 73, 255, 1
		                                            ;MouseClick, Left, 30, 227, 1
		                                            Sleep, 300
		                                            SendRaw, {caso.Identificacion}
		                                            Sleep, 300

		                                            If !FileExist(""C:\Desarrollos-BI\ImagenesRobots\img_BuscarCUN.PNG"")
		                                            {{
			                                            ;MsgBox, No existe Imagen img_BuscarCUN
			                                            throw Exception(""Image Doesn't Exist: img_BuscarCUN.PNG"")
		                                            }}

		                                            Loop, 120
		                                            {{
			                                            WinActivate, ""AC Administración de Clientes""
			                                            Sleep, 500

			                                            CoordMode, Pixel, Window
			                                            ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_BuscarCUN.PNG
			                                            If ErrorLevel = 0
			                                            {{
				                                            ;MsgBox, encuentra Imagen img_BuscarCUN
				                                            Click, %FoundX%, %FoundY% Left, 1
				                                            Sleep, 2000
				                                            Break
			                                            }}
			                                            Sleep, 2000
		                                            }}

		                                            If ErrorLevel != 0
		                                            {{
			                                            ;MsgBox, no funciona la imagen img_BuscarCUN
			                                            throw Exception(""Image Not Found: img_BuscarCUN"")
		                                            }}
                                                    Sleep, 5000     

                                                    ;comprobar cuenta
                                                    Img := RutaImagenGlobal . ""No_Existe_Informacion.PNG""
                                                    ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %Img%
                                                    if (ErrorLevel = 0) {{
                                                        ; imagen encontrada -> devolver estado y no continuar
                                                        result := ""No_Existe_Informacion""
                                                        ;MsgBox, No_Existe_Informacion
                                                        return
                                                    }} 

		                                            Sleep, 3000
		                                            ;~ Damos click en ""Cuenta"" Primera fila
		                                            MouseMove, 390, 135
		                                            Sleep, 600
		                                            MouseClick, Left, 390, 135, 1
		                                            Sleep, 5000

                                                    ; Comprobar si la linea es Prepago
                                                    Imagen := RutaImagenGlobal . ""Prepago.PNG""
                                                    ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %Imagen%
                                                    if (ErrorLevel = 0)
                                                    {{
                                                        ; Imagen encontrada: establecer la bandera y CONTINUAR el script
                                                        lineaEncontrada := 1
    
                                                        ;MsgBox, Línea de Prepago detectada. Continuando...
                                                    }}
                                                    Sleep, 3000

                                                    if (lineaEncontrada = 0)
                                                    {{
                                                        ; Comprobar si la linea es Postpago
                                                        Imagen := RutaImagenGlobal . ""Postpago.PNG""
                                                        ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %Imagen%
                                                        if (ErrorLevel = 0)
                                                        {{
                                                            ; Imagen encontrada: establecer la bandera y CONTINUAR el script
                                                            lineaEncontrada := 1
                                                            ;MsgBox, Línea de Postpago detectada. Continuando...
                                                        }}
                                                    }}

                                                    Sleep, 3000

                                                    if (lineaEncontrada = 0)
                                                    {{
                                                        ; Si la bandera sigue en 0, ninguna de las dos imágenes fue encontrada.
                                                        result := ""FALLO: Tipo de línea no detectado (ni Prepago ni Postpago).""
                                                        ;MsgBox, Fallo: No se detectó ni Prepago ni Postpago. Saliendo del script.
                                                        return ; Detiene la ejecución del script aquí y retorna
                                                    }}


                                                    ;Comprobar si la linea esta desactivada
                                                    Img2 := RutaImagenGlobal . ""Linea_Desactivada.PNG""
                                                    ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %Img2%
                                                    if (ErrorLevel = 0) {{
                                                        ; imagen encontrada -> devolver estado y no continuar
                                                        result := ""Linea_Desactivada""
                                                        ;MsgBox, Linea_Desactivada
                                                        return
                                                    }} 
                                                    Sleep, 3000

		                                            ;~ Damos click en ""Cuenta"" Primera fila
		                                            MouseMove, 390, 305
		                                            Sleep, 600
		                                            MouseClick, Left, 390, 305, 1
		                                            Sleep, 3000

		                                            If !FileExist(""C:\Desarrollos-BI\ImagenesRobots\img_BotonConsultarCUN.PNG"")
		                                            {{
			                                            throw Exception(""Image Doesn't Exist: img_BotonConsultarCUN.PNG"")
		                                            }}

		                                            Loop, 120
		                                            {{
			                                            WinActivate, ""AC Administración de Clientes""
			                                            Sleep, 500

			                                            CoordMode, Pixel, Window
			                                            ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_BotonConsultarCUN.PNG
			                                            If ErrorLevel = 0
			                                            {{
				                                            ;MsgBox, encuentra Imagen img_BotonConsultarCUN
				                                            Click, %FoundX%, %FoundY% Left, 1
				                                            Sleep, 4000
				                                            Break
			                                            }}
			                                            Sleep, 2000
		                                            }}

		                                            If ErrorLevel != 0
		                                            {{
			                                            throw Exception(""Image Not Found: img_BotonConsultarCUN"")
		                                            }}

		                                            Sleep, 5000

		                                            ;If !FileExist(""C:\Desarrollos-BI\ImagenesRobots\img_AceptarLineaProceso.PNG"")
		                                            ;{{
			                                            ;throw Exception(""Image Doesn't Exist: img_AceptarLineaProceso.PNG"")
		                                            ;}}
	
		
		                                            ;Loop, 120
		                                            ;{{
			                                            ;WinActivate, ""AC Administración de Clientes""
			                                            ;Sleep, 500

			                                            ;CoordMode, Pixel, Window
			                                            ;ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_AceptarLineaProceso.PNG
			                                            ;If ErrorLevel = 0
			                                            ;{{
				                                            ;Click, %FoundX%, %FoundY% Left, 1
				                                            ;Sleep, 2000
				                                            ;Break
			                                            ;}}
			                                            ;else
			                                            ;{{
				                                            ;Break
			                                            ;}}
			                                            ;Sleep, 1000
		                                            ;}}

                                                    Sleep, 2000

                                                    ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\ValidacionBarrio.PNG
			                                        If ErrorLevel = 0
			                                        {{
				                                            ;Click, %FoundX%, %FoundY% Left, 1
                                                            validacion := 1
                                                            Sleep, 2000
                                                            Send, {{Enter}}
				                                            Sleep, 3000  
				                                            
			                                        }}
			                                         
			                                        Sleep, 2000

                                                    if (validacion = 1)
                                                    {{
                                                        ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_BotonAceptarCampoBarrio.PNG
			                                            If ErrorLevel = 0
			                                            {{
				                                                ;Click, %FoundX%, %FoundY% Left, 1
                                                                Send, {{Enter}}
				                                                Sleep, 2000
                                                                Send, {{Tab 36}}
                                                                Sleep, 3000
				                                            
			                                            }}
			                                         
			                                            Sleep, 1000
                                                    }}
                                                    else 
                                                    {{
                                                        ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_BotonAceptarCampoBarrio.PNG
			                                            If ErrorLevel = 0
			                                            {{
				                                                ;Click, %FoundX%, %FoundY% Left, 1
                                                                Send, {{Enter}}
				                                                Sleep, 2000
                                                                Send, {{Tab 37}}
                                                                Sleep, 3000
				                                            
			                                            }}
			                                         
			                                            Sleep, 1000
                                                    }}
			                                         
			                                        Sleep, 2000

                                                    ;quitar mensajes informativos 
		                                              ;por si sale mensaje de confirmacion 
		                                            ;Send, {{Enter}}
                                                    ;Sleep, 3000
                                                    ;Send, {{Enter}}
                                                    ;Sleep, 3000
                                                    ;Send, {{Enter}}		
		                                            ;Sleep, 3000
		                                            
		                                            ;If !FileExist(""C:\Desarrollos-BI\ImagenesRobots\img_CiudadDptoIncorrectos.PNG"")
		                                             ;{{
			                                            ; throw Exception(""Image Doesn't Exist: img_CiudadDptoIncorrectos.PNG"")
		                                             ;}}

		                                             ;Loop, 30
		                                            ; {{
			                                             ;WinActivate, ""Validación Ciudad/Departamento""
			                                             ;Sleep, 500

			                                            CoordMode, Pixel, Window
			                                             ;ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_CiudadDptoIncorrectos.PNG
			                                             ;If ErrorLevel = 0
			                                             ;{{
				                                            ; Click, %FoundX%, %FoundY% Left, 1
				                                            ;  Sleep, 2000
				                                             ;Break
			                                           ;  }}
			                                             ;else
			                                             ;{{
				                                       ;       Break
			                                           ;  }}
			                                            ; Sleep, 1000
		                                             ;}}

		                                            ;If !FileExist(""C:\Desarrollos-BI\ImagenesRobots\img_BotonAceptarFechaCum.PNG"")
		                                            ;{{
			                                         ;   throw Exception(""Image Doesn't Exist: img_BotonAceptarFechaCum.PNG"")
		                                            ;}}

		                                            ;Loop, 30
		                                            ;{{
			                                         ;   WinActivate, ""AC Administración de Clientes""
			                                          ;  Sleep, 500

			                                           ; CoordMode, Pixel, Window
			                                            ;ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_BotonAceptarFechaCum.PNG
			                                            ;If ErrorLevel = 0
			                                            ;{{
				                                         ;   Click, %FoundX%, %FoundY% Left, 1
				                                          ;  Sleep, 2000
				                                           ; Break
			                                            ;}}
			                                            ;else
			                                            ;{{
				                                         ;   Break
			                                           ; }}
			                                            ;Sleep, 1000
		                                           ; }}


		                                           ; If !FileExist(""C:\Desarrollos-BI\ImagenesRobots\img_FalloConexionBD.PNG"")
		                                           ; {{
			                                      ;      throw Exception(""Image Doesn't Exist: img_FalloConexionBD.PNG"")
		                                          ;  }}

		                                           ; Loop, 30
		                                       ;     {{
			                                          ;  WinActivate, ""AC Administración de Clientes""
			                                         ;   Sleep, 500

			                                          ;  CoordMode, Pixel, Window
			                                          ;  ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Desarrollos-BI\ImagenesRobots\img_FalloConexionBD.PNG
			                                          ;  If ErrorLevel = 0
			                                           ; {{
				                                            ;MsgBox, encuentra Imagen img_FalloConexionBD
				                                           ; Send, {{Enter}}
				                                          ;  Sleep, 2000
				                                          ;  Break
			                                          ;  }}
                                                        ;else
			                                          ;  {{
				                                        ;    Break
			                                           ; }}
			                                          ;  Sleep, 1000
		                                           ; }}

		                                            result := ""OK""
		                                            ;MsgBox, Valor variable Resultado %Resultado%

	                                            }}
	                                            catch e
	                                            {{
		                                            result := ""(AHK "" e.What "": "" e.Line "") "" e.Message
	                                            }}

	                                            return
                                                            
                                ";


                                try
                                {
                                    ahk.ExecRaw(Consulta2.ToString());
                                    Thread.Sleep(2000);
                                    result = ahk.GetVar("result");
                                    if (result == "OK") // Asegúrate de que el resultado sea una cadena "true"
                                    {
                                        Console.WriteLine("OK");
                                    }
                                    else if (result.Contains("Image Doesn't Exist"))
                                    {
                                        log.Escribir("Error critico! imagen no existe en la ruta.");
                                        log.Escribir($"Verifique que las imagenes se encuentran en la ruta: C:\\Desarrollos-BI\\ImagenesRobots\\");
                                        Console.WriteLine("Error critico! imagen no existe en la ruta.");
                                        Console.WriteLine($"Verifique que las imagenes se encuentran en la ruta: C:\\Desarrollos-BI\\ImagenesRobots\\");
                                        _emailSender.SendMailException("Error critico! imagen no encontrada", caso?.Id_Lecturabilidad, "Verifique que la imagenes se encuentran en la ruta: C:\\Desarrollos-BI\\ImagenesRobots\\");
                                        log.finalizar();

                                        // Terminar el programa ya que es un error critico
                                        Environment.Exit(1);
                                        //throw new ACImgException($"Error critico: imagen: img_LogoAC.PNG no existe en la ruta.");
                                    }
                                    else if (result.Contains("Image Not Found"))
                                    {
                                        log.Escribir("Error! imagen no encontrada.");
                                        Console.WriteLine("Error! imagen no encontrada.");
                                        throw new ACImgException("Error! imagen no encontrada");
                                    }
                                    else if (result == "")
                                    {
                                        log.Escribir("Error critico! Al ejecutar script ahk");
                                        throw new AHKException($"Error: fallo al ejecutar AHK.");
                                    }
                                    else if (result.Contains("No_Existe_Informacion"))
                                    {
                                        log.Escribir("No existe informacion para la identificacion consultada");
                                        Console.WriteLine("No existe informacion para la identificacion consultada");
                                        _casoRepo.ActualizarEstado(
                                            caso.Id_Lecturabilidad,
                                            nuevoEstado: "RECHAZADO",
                                            observaciones: $"No existe informacion para la identificacion consultada",
                                            ExtraccionCompleta: caso.ExtraccionCompleta
                                        );

                                        goto BuscarCaso;
                                    }
                                    else if (result.Contains("Linea_Desactivada"))
                                    {
                                        log.Escribir("La linea asociada a la identificacion consultada se encuentra desactivada");
                                        Console.WriteLine("La linea asociada a la identificacion consultada se encuentra desactivada");
                                        _casoRepo.ActualizarEstado(
                                            caso.Id_Lecturabilidad,
                                            nuevoEstado: "RECHAZADO",
                                            observaciones: $"La linea asociada a la identificacion consultada se encuentra desactivada",
                                            ExtraccionCompleta: caso.ExtraccionCompleta
                                        );
                                        goto BuscarCaso;
                                    }
                                    else if (result.Contains("No_Tiene_Linea"))
                                    {
                                        log.Escribir("La identificacion consultada no tiene linea asociada");
                                        Console.WriteLine("La identificacion consultada no tiene linea asociada");
                                        _casoRepo.ActualizarEstado(
                                            caso.Id_Lecturabilidad,
                                            nuevoEstado: "RECHAZADO",
                                            observaciones: $"La identificacion consultada no tiene linea asociada",
                                            ExtraccionCompleta: caso.ExtraccionCompleta
                                        );
                                        goto BuscarCaso;
                                    }
                                    else if (result.Contains("FALLO: Tipo de línea no detectado (ni Prepago ni Postpago)."))
                                    {
                                        log.Escribir("La línea asociada a la identifición no tiene tipo de línea (Prepago/Postpago).");
                                        Console.WriteLine("La línea asociada a la identifición no tiene tipo de línea (Prepago/Postpago).");
                                        _casoRepo.ActualizarEstado(
                                            caso.Id_Lecturabilidad,
                                            nuevoEstado: "RECHAZADO",
                                            observaciones: $"Sin linea activa",
                                            ExtraccionCompleta: caso.ExtraccionCompleta
                                        );
                                        goto BuscarCaso;
                                    }
                                    else
                                    {
                                        throw new ACException($"Error: En AC.");
                                    }

                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error al ejecutar el script: {ex.Message}");
                                    throw new ACException($"Error: En AC.", ex);
                                }

                                Thread.Sleep(2000);
                                string ventana2 = @"
                                            SetTitleMatchMode, 2

                                            WinTitle := ""AC Administración de Clientes""

                                            ; Intentar activar y traer al frente la ventana
                                            WinActivate, %WinTitle%
                                            WinWaitActive, %WinTitle%, , 5
	                                
                                            ; Forzar Foreground
                                            WinExistId := WinExist(WinTitle)
                                            if (WinExistId)
                                            {
                                                ; asegurar foreground (DLL)
                                                DllCall(""SetForegroundWindow"", ""ptr"", WinExistId)
                                                ; temporalmente arriba
                                                WinSet, AlwaysOnTop, On, ahk_id %WinExistId%
                                            }
                                            Sleep, 300
                                            ";

                                try
                                {
                                    ahk.ExecRaw(ventana2.ToString());
                                    Thread.Sleep(2000);
                                    result = ahk.GetVar("result");
                                    if (result == "OK") // Asegúrate de que el resultado sea una cadena "true"
                                    {
                                        Console.WriteLine("OK");
                                    }
                                    else
                                    {
                                        throw new ACException($"Error: En AC.");
                                    }

                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error al ejecutar el script: {ex.Message}");
                                    throw new ACException($"Error: En AC.", ex);
                                }

                                if (string.IsNullOrEmpty(caso.DirigidoA) || caso.DirigidoA == "0")
                                {
                                    // Captura de nombre
                                    Thread.Sleep(3000);
                                    string Nombre2 = @"
                                            global result
                                            global nombre
                                            global apellido
                                            nombre := """"
                                            apellido := """"
                                            SetTitleMatchMode, 2

                                            WinActivate, ""AC Administración de Clientes""
		                                    Sleep, 1200

                                            ;usar coordenadas relativas a la ventana
                                            CoordMode, Mouse, Window
                                            CoordMode, Pixel, Window

                                            ; posicionar sobre la celda de la Nombre y copiar
                                            ;MouseMove, 278, 164
                                            Sleep, 500
                                            ;Click, left, 278, 164
                                            Send, {Tab}
                                            Sleep, 1000
                                            Send, +{Right 40}
                                            Sleep, 2000
                                            Clipboard := """"               ; limpiar clipboard
                                            SendInput, ^c
                                            ClipWait, 2                 ; espera hasta 2s (ajusta)
                                            if ErrorLevel
                                            {{
                                                ; no copió: retry pequeño y volver a intentar una copia directa
                                                Sleep, 150
                                                SendInput, ^c
                                                ClipWait, 1
                                            }}

                                            current := RegExReplace(Clipboard, ""^\s+|\s+$"", """") ; trim
                                            ;MsgBox, copiado  %current%
                                            nombre := current

                                            ; posicionar sobre la celda de la Apellido y copiar
                                            ;MouseMove, 278, 189
                                            Sleep, 500
                                            ;Click, left, 278, 189
                                            Send, {Tab}
                                            Sleep, 1000
                                            Send, +{Right 40}
                                            Sleep, 2000
                                            Clipboard := """"               ; limpiar clipboard
                                            SendInput, ^c
                                            ClipWait, 2                 ; espera hasta 2s (ajusta)
                                            if ErrorLevel
                                            {{
                                                ; no copió: retry pequeño y volver a intentar una copia directa
                                                Sleep, 150
                                                SendInput, ^c
                                                ClipWait, 1
                                            }}

                                            current := RegExReplace(Clipboard, ""^\s+|\s+$"", """") ; trim
                                            ;MsgBox, copiado %current%
                                            apellido := current

                                            ; Normalizar variables (evitar pipes en los valores)
                                            if (nombre = """")
                                                nombre := """"
                                            else
                                                nombre := StrReplace(nombre, ""|"", ""-"")  ; por seguridad, quitar pipes

                                            if (apellido = """")
                                                apellido := """"
                                            else
                                                apellido := StrReplace(apellido, ""|"", ""-"")

                                            if (nombre != """" && apellido != """")
                                            {{
                                                result := ""OK|"" . nombre . ""|"" . apellido
                                                return
                                            }}
                                            else
                                            {{
                                                result := ""FAIL_NOTFOUND""
                                                return
                                            }}
                                        
                                        ";

                                    try
                                    {
                                        ahk.ExecRaw(Nombre2.ToString());
                                        Thread.Sleep(2000);
                                        result = ahk.GetVar("result");
                                        tabular = 1;
                                        if (result.Contains("OK")) // Asegúrate de que el resultado sea una cadena "true"
                                        {
                                            Console.WriteLine("OK");
                                            log.Escribir("Nombre extraido...");

                                            // Separar por '|'
                                            var parts = result.Split('|');
                                            status = parts.Length > 0 ? parts[0] : "";
                                            nombre = parts.Length > 1 ? parts[1] : "";
                                            apellido = parts.Length > 2 ? parts[2] : "";
                                            nombreCompleto = nombre + " " + apellido;
                                            // Guardar en DTO/objeto
                                            caso.DirigidoA = nombreCompleto;

                                            // Persistir identificacion 
                                            try
                                            {
                                                if (!string.IsNullOrEmpty(caso.DirigidoA) && caso.DirigidoA != "0")
                                                {
                                                    _casoRepo.ActualizarDatos(caso.Id_Lecturabilidad, "DirigidoA", caso.DirigidoA);

                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                log.Escribir($"Error al insertar en base de datos, nota RR para SolicitudId={caso.Id_Lecturabilidad}: {ex.Message}");
                                                Console.WriteLine($"Error al insertar en base de datos, nota RR para SolicitudId={caso.Id_Lecturabilidad}: {ex.Message}");
                                                throw;
                                            }
                                        }
                                        else if (result.Contains("FAIL_NOTFOUND"))
                                        {
                                            log.Escribir("No se pudo extraer el nombre");
                                            Console.WriteLine("No se pudo extraer el nombre");
                                            throw new ACException($"Error: En AC.");
                                        }
                                        else if (result == "")
                                        {
                                            log.Escribir("Error critico! Al ejecutar script ahk");
                                            throw new AHKException($"Error: fallo al ejecutar AHK.");
                                        }
                                        else
                                        {
                                            throw new ACException($"Error: En AC.");
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Error al ejecutar el script: {ex.Message}");
                                        throw new ACException($"Error: En AC.");
                                    }


                                }

                                if (string.IsNullOrEmpty(caso.CorreoaNotificar) || caso.CorreoaNotificar == "0")
                                {
                                    // Captura de correo
                                    Thread.Sleep(3000);
                                    string scriptcorreo2 = $@"
                                            global result
                                            global correo
                                            global tabular 
                                            tabular := ""{tabular}""
                                            correo := """"
                                            SetTitleMatchMode, 2

                                            WinActivate, ""AC Administración de Clientes""
		                                    Sleep, 1200

                                            ;usar coordenadas relativas a la ventana
                                            CoordMode, Mouse, Window
                                            CoordMode, Pixel, Window

                                            ; posicionar sobre la celda de la Nombre y copiar
                                            ;MouseMove, 277, 322
                                            Sleep, 500
                                            ;MouseClick, left, 277, 322
                                            if (tabular = 0)
                                            {{
                                                Send, {{Tab 8}}
                                            }}
                                            else if (tabular = 1)
                                            {{
                                                Send, {{Tab 6}}
                                            }}

                                            Sleep, 1000
                                            Send, +{{Right 20}}
                                            Sleep, 2000
                                            Clipboard := """"               ; limpiar clipboard
                                            Send, ^c
                                            ClipWait, 2                 ; espera hasta 2s (ajusta)
                                            if ErrorLevel
                                            {{
                                                ; no copió: retry pequeño y volver a intentar una copia directa
                                                Sleep, 150
                                                Send, ^c
                                                ClipWait, 1
                                            }}

                                            current := RegExReplace(Clipboard, ""^\s+|\s+$"", """") ; trim
                                            ;MsgBox, copiado pqr %current%
                                            correo := current

                                            ; Normalizar variables (evitar pipes en los valores)
                                            if (correo = """")
                                                correo := """"
                                            else
                                                correo := StrReplace(correo, ""|"", ""-"")  ; por seguridad, quitar pipes

                                            if (correo != """")
                                            {{
                                                result := ""OK|"" . correo
                                                return
                                            }}
                                            else
                                            {{
                                                result := ""FAIL_NOTFOUND""
                                                return
                                            }}
                                        
                                        ";

                                    try
                                    {
                                        ahk.ExecRaw(scriptcorreo2.ToString());
                                        Thread.Sleep(2000);
                                        result = ahk.GetVar("result");
                                        tabular = 2;
                                        if (result.Contains("OK")) // Asegúrate de que el resultado sea una cadena "true"
                                        {
                                            Console.WriteLine("OK");
                                            log.Escribir("correo extraido...");

                                            // Separar por '|'
                                            var parts = result.Split('|');
                                            status = parts.Length > 0 ? parts[0] : "";
                                            correo = parts.Length > 1 ? parts[1] : "";

                                            // Guardar en DTO/objeto
                                            caso.CorreoaNotificar = correo;

                                            // Persistir la correo 
                                            try
                                            {
                                                if (!string.IsNullOrEmpty(caso.CorreoaNotificar) && caso.CorreoaNotificar != "0")
                                                {
                                                    _casoRepo.ActualizarDatos(caso.Id_Lecturabilidad, "CorreoaNotificar", caso.CorreoaNotificar);
                                                    //goto BuscarNotasHistorico;

                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                log.Escribir($"Error al insertar en base de datos, nota RR para SolicitudId={caso.Id_Lecturabilidad}: {ex.Message}");
                                                Console.WriteLine($"Error al insertar en base de datos, nota RR para SolicitudId={caso.Id_Lecturabilidad}: {ex.Message}");
                                                throw;
                                            }
                                        }
                                        else if (result.Contains("FAIL_NOTFOUND"))
                                        {
                                            log.Escribir("No se pudo extraer la identificacion");
                                            Console.WriteLine("No se pudo extraer la identificacion");
                                            throw new ACException($"Error: En AC.");

                                        }
                                        else if (result == "")
                                        {
                                            log.Escribir("Error critico! Al ejecutar script ahk");
                                            throw new AHKException($"Error: fallo al ejecutar AHK.");
                                        }
                                        else
                                        {
                                            throw new ACException($"Error: En AC.");
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Error al ejecutar el script: {ex.Message}");
                                        throw new ACException($"Error: En AC.");
                                    }
                                }

                                if (string.IsNullOrEmpty(caso.Cuscode) || caso.Cuscode == "0")
                                {

                                    // Captura de Cuscode
                                    Thread.Sleep(3000);
                                    string scriptCuscode = $@"
                                            global result
                                            global cuscode
                                            global tabular
                                            tabular := ""{tabular}""
                                            cuscode := """"
                                            SetTitleMatchMode, 2

                                            WinActivate, ""AC Administración de Clientes""
		                                    Sleep, 1200

                                            ;usar coordenadas relativas a la ventana
                                            CoordMode, Mouse, Window
                                            CoordMode, Pixel, Window

                                            ; posicionar sobre la celda de la Nombre y copiar
                                            ;MouseMove, 244, 92
                                            Sleep, 500
                                            if (tabular = 0)
                                            {{
                                                Send, {{Tab 38}}
                                            }}
                                            else if (tabular = 1)
                                            {{
                                                Send, {{Tab 36}}
                                            }}
                                            else if (tabular = 2)
                                            {{
                                                Send, {{Tab 30}}
                                            }}
                                            Sleep, 500
                                            Send, {{Right 3}}
                                            Sleep, 500
                                            Clipboard := """"               ; limpiar clipboard
                                            Sleep, 1000
                                            Send, {{Left 2}}
                                            ;MouseClick, left, 244, 92, 1
                                            Sleep, 1000
                                            Send, +{{Right 20}}
                                            Sleep, 1000
                                            Send, ^c
                                            Sleep, 500
                                            ClipWait, 2                 ; espera hasta 2s (ajusta)

                                            if ErrorLevel
                                            {{
                                                ; no copió: retry pequeño y volver a intentar una copia directa
                                                Sleep, 150
                                                Send, ^c
                                                ClipWait, 1
                                            }}

                                            current := RegExReplace(Clipboard, ""^\s+|\s+$"", """") ; trim
                                            ;MsgBox, copiado %current%
                                            cuscode := current

                                            ; Normalizar variables (evitar pipes en los valores)
                                            if (cuscode = """")
                                                cuscode := """"
                                            else
                                                cuscode := StrReplace(cuscode, ""|"", ""-"")  ; por seguridad, quitar pipes

                                            if (cuscode != """")
                                            {{
                                                result := ""OK|"" . cuscode
                                                return
                                            }}
                                            else
                                            {{
                                                result := ""FAIL_NOTFOUND""
                                                return
                                            }}
                                        
                                        ";

                                    try
                                    {
                                        ahk.ExecRaw(scriptCuscode.ToString());
                                        Thread.Sleep(2000);
                                        result = ahk.GetVar("result");

                                        if (result.Contains("OK")) // Asegúrate de que el resultado sea una cadena "true"
                                        {
                                            Console.WriteLine("OK");
                                            log.Escribir("cuscode extraido...");

                                            // Separar por '|'
                                            var parts = result.Split('|');
                                            status = parts.Length > 0 ? parts[0] : "";
                                            custcode = parts.Length > 1 ? parts[1] : "";

                                            // Guardar en DTO/objeto
                                            caso.Cuscode = custcode;

                                            if (!string.IsNullOrEmpty(caso.Cuscode) && caso.Cuscode != "" && caso.Cuscode != "0")
                                            {
                                                CuscodeValido = ACHelper.IsValidCuscode(caso.Cuscode);
                                            }

                                            if (!CuscodeValido)
                                            {
                                                log.Escribir("Cuscode extraido no es valido.");
                                                Console.WriteLine("Cuscode extraido no es valido.");
                                                throw new ACException($"Error: Cuscode no es valido.");
                                            }

                                            // Persistir Cuscode 
                                            try
                                            {
                                                if (!string.IsNullOrEmpty(caso.Cuscode) && caso.Cuscode != "0")
                                                {
                                                    _casoRepo.ActualizarDatos(caso.Id_Lecturabilidad, "Cuscode", caso.Cuscode);

                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                log.Escribir($"Error al insertar en base de datos, nota RR para SolicitudId={caso.Id_Lecturabilidad}: {ex.Message}");
                                                Console.WriteLine($"Error al insertar en base de datos, nota RR para SolicitudId={caso.Id_Lecturabilidad}: {ex.Message}");
                                                throw;
                                            }
                                        }
                                        else if (result.Contains("FAIL_NOTFOUND"))
                                        {
                                            log.Escribir("No se pudo extraer la Cuscode");
                                            Console.WriteLine("No se pudo extraer la Cuscode");
                                            throw new ACException($"Error: En AC.");

                                        }
                                        else if (result.Contains("FAIL_NOTFOUND"))
                                        {
                                            log.Escribir("No se pudo extraer la identificacion");
                                            Console.WriteLine("No se pudo extraer la identificacion");
                                            throw new ACException($"Error: En AC.");

                                        }
                                        else if (result == "")
                                        {
                                            log.Escribir("Error critico! Al ejecutar script ahk");
                                            throw new AHKException($"Error: fallo al ejecutar AHK.");
                                        }
                                        else
                                        {
                                            throw new ACException($"Error: En AC.");
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Error al ejecutar el script: {ex.Message}");
                                        throw new ACException($"Error: En AC.");
                                    }
                                }

                                if (caso.DirigidoA != "0" && caso.Cuscode != "0" && caso.CorreoaNotificar != "0")
                                {
                                    log.Escribir("Datos extraidos");
                                    Console.WriteLine("Datos extraidos");
                                    caso.ExtraccionCompleta = true;
                                    goto TerminarIntento;
                                }
                                else
                                {
                                    log.Escribir("Datos no pudieron ser extraidos");
                                    Console.WriteLine("Datos no pudieron ser extraidos");
                                    caso.ExtraccionCompleta = false;
                                    goto TerminarIntento;
                                }

                            }
                            else
                            {
                                //Devolver por datos incompletos
                                _casoRepo.ActualizarEstado(
                                    caso.Id_Lecturabilidad,
                                    nuevoEstado: "RECHAZADO",
                                    observaciones: $"Caso Rechazado por datos incompletos",
                                    ExtraccionCompleta: caso.ExtraccionCompleta
                                );

                                goto BuscarCaso;
                            }

                        TerminarIntento:

                            // Marcar como completo
                            if (caso.ExtraccionCompleta)
                            {
                                log.Escribir("Extracción completa");
                                _casoRepo.ActualizarEstado(
                                    caso.Id_Lecturabilidad,
                                    nuevoEstado: "EXITOSO",
                                    observaciones: $"Extracción completa",
                                    ExtraccionCompleta: caso.ExtraccionCompleta
                                );
                                log.Escribir($"Caso: {caso.Id_Lecturabilidad} marcado como EXITOSO.");
                                Thread.Sleep(2000);

                                // Obtener el remitente a comprobar 
                                remitente = (caso.CorreoaNotificar ?? string.Empty).Trim();

                                // Normalizar: bajar a minúsculas y eliminar espacios laterales
                                remitenteNorm = remitente.ToLowerInvariant();

                                // Si el remitente está en la lista -> NO insertar
                                if (!string.IsNullOrEmpty(remitente) && blockedSenders.Contains(remitenteNorm))
                                {
                                    log.Escribir($"Remitente bloqueado detectado para Id {caso.Id_Lecturabilidad}: {remitente} -> No se insertará en DetalleLecturabilidad. Atención manual requerida.");
                                    Console.WriteLine($"Remitente bloqueado ({remitente}) — se omite InsertarDetalle para Id {caso.Id_Lecturabilidad}.");

                                    // marcar el caso para revisión manual en lugar de dejarlo como EXITOSO.

                                    _casoRepo.ActualizarEstado(
                                        caso.Id_Lecturabilidad,
                                        nuevoEstado: "RECHAZADO",
                                        observaciones: $"Remitente bloqueado ({remitente}). Atención manual requerida.",
                                        ExtraccionCompleta: caso.ExtraccionCompleta
                                    );
                                    log.Escribir($"Caso {caso.Id_Lecturabilidad} marcado como PENDIENTE_MANUAL por remitente bloqueado.");


                                    // No llamamos a InsertarDetalle; salimos del flujo de inserción.
                                }
                                else 
                                {
                                    // remitente NO bloqueado -> proceder con la inserción
                                    try
                                    {
                                        int filas = _actualizarCasoRepo.InsertarDetalle(caso.Id_Lecturabilidad);
                                        Console.WriteLine($"InsertarDetalle: filas afectadas = {filas}");
                                        log.Escribir($"InsertarDetalle: filas afectadas = {filas}");
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("Fallo InsertarDetalle: " + ex.Message);
                                        log.Escribir("Fallo InsertarDetalle: " + ex.Message);
                                        throw;
                                    }
                                }

                                tabular = 0; //reseteo
                                procesoCompleto = true; // Marcar como completo para salir del bucle
                            }
                            else
                            {
                                //Marcar caso gestionado pero sin notas
                                _casoRepo.ActualizarEstado(
                                    caso.Id_Lecturabilidad,
                                    nuevoEstado: "FALLO_RPA",
                                    observaciones: $"Error en la extraccion de datos",
                                    ExtraccionCompleta: caso.ExtraccionCompleta
                                );
                                log.Escribir($"Error en la extraccion de datos. Caso {caso.Id_Lecturabilidad}.");
                                Console.WriteLine($"Error en la extraccion de datos. Caso {caso.Id_Lecturabilidad}");
                                Thread.Sleep(2000);
                                tabular = 0; //reseteo
                                procesoCompleto = true; // Marcar como completo para salir del bucle
                            }
                        }
                        catch (ACImgException ex)
                        {
                            intentos++;
                            _casoRepo.IncrementarIntento(caso.Id_Lecturabilidad);
                            log.Escribir("Error de busqueda de imagen");

                            if (intentos < _maxLoginAttempts)
                            {
                                log.Escribir($"Reintentando caso {caso.Id_Lecturabilidad}");
                            }
                            else
                            {
                                log.Escribir("Fallo tras 3 intentos");
                                _casoRepo.ActualizarEstado(
                                    caso.Id_Lecturabilidad,
                                    nuevoEstado: "FALLO_RPA",
                                    observaciones: "3 intentos fallidos",
                                    ExtraccionCompleta: caso.ExtraccionCompleta
                                );
                                tabular = 0; //reseteo
                                procesoCompleto = true;
                                goto BuscarCaso;

                            }
                        }
                        catch (AHKException)
                        {
                            intentos++;
                            _casoRepo.IncrementarIntento(caso.Id_Lecturabilidad);
                            log.Escribir("Error de ahk");

                            if (intentos < _maxLoginAttempts)
                            {
                                log.Escribir($"Reintentando caso {caso.Id_Lecturabilidad}");
                            }
                            else
                            {
                                log.Escribir("Fallo tras 3 intentos");
                                _casoRepo.ActualizarEstado(
                                    caso.Id_Lecturabilidad,
                                    nuevoEstado: "FALLO_RPA",
                                    observaciones: "3 intentos fallidos",
                                    ExtraccionCompleta: caso.ExtraccionCompleta
                                );
                                tabular = 0; //reseteo
                                procesoCompleto = true;
                                goto BuscarCaso;

                            }
                        }
                        catch (ACException ex)
                        {
                            intentos++;
                            _casoRepo.IncrementarIntento(caso.Id_Lecturabilidad);
                            log.Escribir("Error en AC");

                            if (intentos < _maxLoginAttempts)
                            {
                                log.Escribir($"Reintentando caso {caso.Id_Lecturabilidad}");
                            }
                            else
                            {
                                log.Escribir("Fallo en AC tras 3 intentos");
                                _casoRepo.ActualizarEstado(
                                    caso.Id_Lecturabilidad,
                                    nuevoEstado: "FALLO_RPA",
                                    observaciones: "3 intentos fallidos",
                                    ExtraccionCompleta: caso.ExtraccionCompleta
                                );
                                tabular = 0; //reseteo
                                procesoCompleto = true;
                                goto BuscarCaso;

                            }

                        }

                    }


                }
                catch (Exception ex)
                {
                    _casoRepo.IncrementarIntento(caso.Id_Lecturabilidad);

                    _casoRepo.ActualizarEstado(
                                        caso.Id_Lecturabilidad,
                                        nuevoEstado: "FALLO_RPA",
                                        observaciones: $"Error inesperado en el proceso...",
                                        ExtraccionCompleta: caso.ExtraccionCompleta
                                    );

                    // Solo para errores inesperados (no relacionados con ACR)
                    log.Escribir($"[Error crítico inesperado]: {ex.Message}");
                    Console.WriteLine($"[Error crítico inesperado]: {ex.Message}");
                    _emailSender.SendMailException("Error crítico inesperado:" + ex.Message, caso.Id_Lecturabilidad, ex.Message.ToString());                   

                    log.Escribir("Finalizando ejecución del RPA de validacion de datos AC...");
                    log.finalizar();                 

                    // Terminar el programa ya que es un error no controlado
                    Environment.Exit(1);
                }

            }

        }
    
    }
}
