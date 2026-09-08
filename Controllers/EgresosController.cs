using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using TPI_GESTION_HOGAR.Datos;
using TPI_GESTION_HOGAR.Models;

namespace TPI_GESTION_HOGAR.Controllers
{
    public class EgresosController : Controller
    {
        private readonly AppDbContext _context;

        public EgresosController(AppDbContext context)
        {
            _context = context;
        }

       
        [HttpGet]
        public async Task<IActionResult> Crear(int registroId)
        {
               var registro = await _context.Registros
                .Include(r => r.Mujer)
                .FirstOrDefaultAsync(r => r.Id == registroId);

            if (registro == null) return NotFound();

            ViewBag.NombreResidente = $"{registro.Mujer?.Apellido}, {registro.Mujer?.Nombre}";
            ViewBag.RegistroId = registroId;

            var nuevoEgreso = new Egreso
            {
                RegistroId = registroId,
                Fecha = DateOnly.FromDateTime(DateTime.Now) 
            };

            return View(nuevoEgreso);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        
        public async Task<IActionResult> Crear(Egreso egreso, bool seRetiraSola)
        {
            // 1. Limpiamos los campos del referente si se retira sola
            if (seRetiraSola)
            {
                egreso.ApellidoRef = null;
                egreso.NombreRef = null;
                egreso.DNIRef = null;
            }

            ModelState.Remove("Registro");

            if (ModelState.IsValid)
            {
                // 2. Agregamos el registro de egreso con la FECHA ELEGIDA EN LA VISTA
                _context.Egresos.Add(egreso);

                // 3. Buscamos el registro de ingreso activo para cerrarlo
                var registro = await _context.Registros
                    .Include(r => r.Mujer)
                    .FirstOrDefaultAsync(r => r.Id == egreso.RegistroId);

                if (registro != null)
                {
                    // Cerramos el registro actual (Ingreso finalizado)
                    registro.Estado = false;

                    // Liberamos la habitación desasignándola del registro
                    registro.HabitacionId = null;

                    if (registro.Mujer != null)
                    {
                        // Pasamos a la residente a estado inactiva (Historial)
                        registro.Mujer.Estado = false;
                    }
                }

                // 4. Guardamos todo en la base de datos
                await _context.SaveChangesAsync();

                TempData["MensajeExito"] = $"Egreso registrado correctamente con fecha {egreso.Fecha:dd/MM/yyyy}. La cama ha sido liberada.";

                return RedirectToAction("Index", "Home");
            }

            return View(egreso);
        }
    }
}
