using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using TPI_GESTION_HOGAR.Datos;
using TPI_GESTION_HOGAR.Models;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using System.IO;


namespace TPI_GESTION_HOGAR.Controllers
{
    public class RegistrosController : Controller
    {
        private readonly AppDbContext _context;

        public RegistrosController(AppDbContext context)
        {
            _context = context;
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ImportarHistorialBasico(IFormFile archivoExcel)
        {
            if (archivoExcel == null || archivoExcel.Length == 0)
            {
                TempData["MensajeError"] = "Por favor, seleccione un archivo Excel válido.";
                return RedirectToAction("Index", "Home");
            }

            using (var stream = new MemoryStream())
            {
                await archivoExcel.CopyToAsync(stream);
                using (var workbook = new XLWorkbook(stream))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RangeUsed().RowsUsed().Skip(1);

                    int casosImportados = 0;
                    int casosFallidos = 0;

                    using (var transaction = await _context.Database.BeginTransactionAsync())
                    {
                        try
                        {
                            foreach (var row in rows)
                            {
                                // Envolvemos TODA la lectura de la fila en un try-catch.
                                // Si algo explota (como el OleAut date), saltamos a la siguiente fila sin abortar.
                                try
                                {
                                    // 1. DNI
                                    string dniStr = row.Cell(4).CachedValue.ToString().Replace(".", "").Trim();
                                    if (string.IsNullOrEmpty(dniStr)) dniStr = row.Cell(4).GetString().Replace(".", "").Trim();

                                    if (!int.TryParse(dniStr, out int dniLimpio)) continue;

                                    var mujerExistente = await _context.Mujeres.FirstOrDefaultAsync(m => m.DNI == dniLimpio);
                                    int mujerId;

                                    if (mujerExistente == null)
                                    {
                                        // 2. NOMBRE Y APELLIDO
                                        string nombreCompleto = "";
                                        try { nombreCompleto = row.Cell(2).GetString().Trim(); } catch { }

                                        string nombre = "Sin Nombre";
                                        string apellido = "Desconocido";

                                        if (!string.IsNullOrEmpty(nombreCompleto))
                                        {
                                            if (nombreCompleto.Contains(","))
                                            {
                                                var partes = nombreCompleto.Split(',');
                                                apellido = partes[0].Trim();
                                                nombre = partes.Length > 1 ? partes[1].Trim() : "";
                                            }
                                            else
                                            {
                                                var partes = nombreCompleto.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                                                apellido = partes.Length > 0 ? partes[0] : "";
                                                nombre = partes.Length > 1 ? string.Join(" ", partes.Skip(1)) : "";
                                            }
                                        }

                                        // 3. CIUDAD
                                        string ciudad = "";
                                        try { ciudad = row.Cell(3).GetString().Trim(); } catch { }
                                        if (string.IsNullOrEmpty(ciudad)) ciudad = "No especificada";

                                        var nuevaMujer = new Mujer
                                        {
                                            DNI = dniLimpio,
                                            Apellido = apellido,
                                            Nombre = nombre,
                                            Localidad = ciudad,
                                            Nacionalidad = "Argentina",
                                            FechaNac = new DateOnly(1990, 1, 1),
                                            Estado = false
                                        };

                                        _context.Mujeres.Add(nuevaMujer);
                                        await _context.SaveChangesAsync();
                                        mujerId = nuevaMujer.ID;
                                    }
                                    else
                                    {
                                        mujerId = mujerExistente.ID;
                                    }

                                    // 4. FECHA DE INGRESO
                                    DateTime fechaIngresoDt = DateTime.Today;
                                    try
                                    {
                                        var celdaIngreso = row.Cell(1);
                                        if (celdaIngreso.DataType == XLDataType.DateTime)
                                        {
                                            fechaIngresoDt = celdaIngreso.GetDateTime();
                                        }
                                        else if (DateTime.TryParse(celdaIngreso.GetString(), out DateTime fechaParseada))
                                        {
                                            fechaIngresoDt = fechaParseada;
                                        }
                                    }
                                    catch { /* Falla silenciosa, queda DateTime.Today */ }

                                    // CREAR REGISTRO DE INGRESO
                                    var nuevoRegistro = new Registro
                                    {
                                        Fecha = DateOnly.FromDateTime(fechaIngresoDt),
                                        Estado = false,
                                        MujerID = mujerId,
                                        HabitacionId = null
                                    };

                                    _context.Registros.Add(nuevoRegistro);
                                    await _context.SaveChangesAsync();

                                    // 5. FECHA DE EGRESO
                                    string celdaEgresoStr = "";
                                    try { celdaEgresoStr = row.Cell(5).GetString().Trim(); } catch { }

                                    if (!string.IsNullOrEmpty(celdaEgresoStr) && celdaEgresoStr.ToLower() != "en curso")
                                    {
                                        DateTime fechaEgresoDt = DateTime.Today;
                                        try
                                        {
                                            var celdaEgreso = row.Cell(5);
                                            if (celdaEgreso.DataType == XLDataType.DateTime)
                                            {
                                                fechaEgresoDt = celdaEgreso.GetDateTime();
                                            }
                                            else if (DateTime.TryParse(celdaEgreso.GetString(), out DateTime fechaParseadaEgreso))
                                            {
                                                fechaEgresoDt = fechaParseadaEgreso;
                                            }
                                        }
                                        catch { /* Ignora y usa fecha de hoy */ }

                                        // CREAR EGRESO
                                        var nuevoEgreso = new Egreso
                                        {
                                            RegistroId = nuevoRegistro.Id,
                                            Fecha = DateOnly.FromDateTime(fechaEgresoDt),
                                            DomicilioRef = "" // El modelo lo requiere
                                        };

                                        _context.Egresos.Add(nuevoEgreso);
                                    }

                                    casosImportados++;
                                }
                                catch (Exception)
                                {
                                    // Si una fila entera falla por una celda irrecuperable, la contamos como fallida y SEGUIMOS con la próxima.
                                    casosFallidos++;
                                    continue;
                                }
                            }

                            await _context.SaveChangesAsync();
                            await transaction.CommitAsync();

                            string msj = $"Se importaron {casosImportados} registros correctamente.";
                            if (casosFallidos > 0) msj += $" Se omitieron {casosFallidos} filas por tener datos corruptos.";

                            TempData["MensajeExito"] = msj;
                        }
                        catch (Exception ex)
                        {
                            await transaction.RollbackAsync();
                            TempData["MensajeError"] = "Error masivo al procesar. Detalle: " + ex.Message;
                        }
                    }
                }
            }

            return RedirectToAction("Index", "Home");
        }


        [HttpGet]
        [HttpPost]
        [ValidateAntiForgeryToken]

      
        public async Task<IActionResult> ImportarHistorial(IFormFile archivoExcel)
        {
            if (archivoExcel == null || archivoExcel.Length == 0)
            {
                TempData["MensajeError"] = "Por favor, seleccione un archivo Excel válido.";
                return RedirectToAction("Index", "Home");
            }

            using (var stream = new MemoryStream())
            {
                await archivoExcel.CopyToAsync(stream);
                using (var workbook = new XLWorkbook(stream))
                {
                    var worksheet = workbook.Worksheet(1); // Lee la primera pestaña ("Ingresos")
                    var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Saltea los encabezados

                    int casosImportados = 0;

                    using (var transaction = await _context.Database.BeginTransactionAsync())
                    {
                        try
                        {
                            foreach (var row in rows)
                            {
                                // 1. DNI: Leemos el valor directo y limpiamos
                                string dniStr = row.Cell(4).Value.ToString().Replace(".", "").Trim();
                                if (!int.TryParse(dniStr, out int dniLimpio)) continue;

                                var mujerExistente = await _context.Mujeres.FirstOrDefaultAsync(m => m.DNI == dniLimpio);
                                int mujerId;

                                if (mujerExistente == null)
                                {
                                    // 2. NOMBRE Y APELLIDO
                                    string nombreCompleto = row.Cell(2).Value.ToString().Trim();
                                    var partesNombre = nombreCompleto.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                                    string apellido = partesNombre.Length > 0 ? partesNombre[0] : "Desconocido";
                                    string nombre = partesNombre.Length > 1 ? string.Join(" ", partesNombre.Skip(1)) : "Sin Nombre";

                                    // 3. EDAD Y FECHA DE NACIMIENTO
                                    string edadStr = row.Cell(5).Value.ToString().Trim();
                                    if (!int.TryParse(edadStr, out int edadCalculada) || edadCalculada < 15 || edadCalculada > 100)
                                    {
                                        edadCalculada = 30; // 30 años por defecto si la celda está vacía o tiene texto raro
                                    }
                                    var fechaNac = new DateOnly(DateTime.Today.Year - edadCalculada, 1, 1);

                                    // 4. CREAR MUJER
                                    var nuevaMujer = new Mujer
                                    {
                                        DNI = dniLimpio,
                                        Apellido = apellido,
                                        Nombre = nombre,
                                        Nacionalidad = "Argentina",
                                        FechaNac = fechaNac,
                                        Telefono = row.Cell(8).Value.ToString(),
                                        Localidad = row.Cell(18).Value.ToString(),
                                        Estado = false // Inactiva (Historial)
                                    };

                                    _context.Mujeres.Add(nuevaMujer);
                                    await _context.SaveChangesAsync();
                                    mujerId = nuevaMujer.ID;
                                }
                                else
                                {
                                    mujerId = mujerExistente.ID;
                                }

                                // 5. FECHA DE INGRESO (Extracción segura)
                                DateTime fechaIngresoDt = DateTime.Today; // Por defecto hoy
                                var celdaFecha = row.Cell(1);
                                    
                                if (celdaFecha.DataType == XLDataType.DateTime)
                                {
                                    fechaIngresoDt = celdaFecha.GetDateTime();
                                }
                                else if (DateTime.TryParse(celdaFecha.Value.ToString(), out DateTime fechaParseada))
                                {
                                    fechaIngresoDt = fechaParseada;
                                }

                                // 6. CREAR REGISTRO
                                var nuevoRegistro = new Registro
                                {
                                    Fecha = DateOnly.FromDateTime(fechaIngresoDt),
                                    Estado = false,
                                    MujerID = mujerId,
                                    HabitacionId = null
                                };

                                _context.Registros.Add(nuevoRegistro);
                                await _context.SaveChangesAsync();
                                // 7. CREAR EGRESO (Si el caso ya terminó)
                                string celdaFechaEgreso = row.Cell(11).Value.ToString().Trim();

                                if (!string.IsNullOrEmpty(celdaFechaEgreso) && celdaFechaEgreso.ToLower() != "en curso")
                                {
                                    DateTime fechaEgresoDt = DateTime.Today;
                                    var celdaEgreso = row.Cell(11);

                                    if (celdaEgreso.DataType == XLDataType.DateTime)
                                    {
                                        fechaEgresoDt = celdaEgreso.GetDateTime();
                                    }
                                    else if (DateTime.TryParse(celdaEgreso.Value.ToString(), out DateTime fechaParseada))
                                    {
                                        fechaEgresoDt = fechaParseada;
                                    }

                                    var nuevoEgreso = new Egreso
                                    {
                                        RegistroId = nuevoRegistro.Id,
                                        Fecha = DateOnly.FromDateTime(fechaEgresoDt),

                                        // Usamos las propiedades reales de tu modelo
                                        DomicilioRef = row.Cell(12).Value.ToString().Trim(),
                                        NombreRef = row.Cell(13).Value.ToString().Trim() // La col 13 es "Referente de egreso"
                                    };

                                    _context.Egresos.Add(nuevoEgreso);
                                }
                                casosImportados++;


                            }


                            await _context.SaveChangesAsync();
                            await transaction.CommitAsync();

                            TempData["MensajeExito"] = $"Se importaron {casosImportados} casos históricos correctamente.";
                        }
                        catch (Exception ex)
                        {
                            await transaction.RollbackAsync();
                            TempData["MensajeError"] = "Error al procesar el archivo. Detalle: " + ex.Message;
                        }
                    }
                }
            }

            return RedirectToAction("Index", "Home");
        }
        public async Task<IActionResult> Detalles(int? id)
        {
            if (id == null) return NotFound();

            var registro = await _context.Registros
                .Include(r => r.Mujer)
                    .ThenInclude(m => m.Hijos)
                    .Include(r => r.Mujer)
                        .ThenInclude(m => m.Condiciones)
                            .ThenInclude(c => c.TipoCondicion)
                 .Include(r => r.Mujer)
                    .ThenInclude(m => m.Condiciones)
                        .ThenInclude(c => c.ObservacionCondicion)
                .Include(r => r.Habitacion)        
                .Include(r => r.Agresores)         
                .Include(r => r.Denuncias)
                .ThenInclude(d => d.Medida)
                .ThenInclude(m => m.TipoMedida)

                .FirstOrDefaultAsync(r => r.Id == id);

            if (registro == null) return NotFound();

            return View(registro);
        }
        [HttpGet]
        public async Task<IActionResult> Legajo(int? id)
        {

            if (id == null) return NotFound();

            var registro = await _context.Registros
                .Include(r => r.Mujer)
                .Include(r => r.Seguimientos)
                    .ThenInclude(s => s.Personal)
                .Include(r => r.Documentos)
                .FirstOrDefaultAsync(m => m.Id == id);

            if (registro == null) return NotFound();
            return View(registro);
        }

        [HttpGet]
        public async Task<IActionResult> AsignarHabitacion(int? id)
        {
            if (id == null) return NotFound();

     
            var registro = await _context.Registros
                .Include(r => r.Mujer)
                    .ThenInclude(m => m.Hijos)
                .FirstOrDefaultAsync(r => r.Id == id);

            if (registro == null) return NotFound();

            
            int camasNecesarias = 1 + (registro.Mujer?.Hijos?.Count ?? 0);
            var habitacionesActivas = await _context.Habitaciones
                .Include(h => h.Registros.Where(r => r.Mujer != null && r.Mujer.Estado == true))
                .Where(h => h.Estado == true)
                .ToListAsync();

            var listaHabitaciones = habitacionesActivas
                .Where(h =>
                 
                    h.Id == registro.HabitacionId ||
                   
                    (!h.Registros.Any() && h.Capacidad >= camasNecesarias)
                )
                .Select(h => new SelectListItem
                {
                    Value = h.Id.ToString(),
                    Text = $"Habitación {h.NroHabitacion} - (Capacidad total: {h.Capacidad} plazas)",
                    Selected = h.Id == registro.HabitacionId
                }).ToList();

            ViewBag.Habitaciones = listaHabitaciones;
            ViewBag.CamasNecesarias = camasNecesarias;

            return View(registro);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> AsignarHabitacion(int id, int? habitacionId)
        {
            var registro = await _context.Registros.FindAsync(id);
            if (registro == null || registro.Id != id) return NotFound();

            // Actualizamos la habitación (Si elige "Sin asignar", habitacionId viaja como null y es válido gracias a tu migración)
            registro.HabitacionId = habitacionId;

            await _context.SaveChangesAsync();

            TempData["MensajeExito"] = "La asignación de cama se actualizó correctamente.";

            // Al terminar, la devolvemos al Dashboard principal
            return RedirectToAction("Index", "Home");
        }

        
    }
}
