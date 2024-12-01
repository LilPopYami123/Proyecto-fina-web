using ManagmentApplication.Data;
using ManagmentApplication.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace ManagmentApplication.Controllers
{
    public class ControladorController : Controller
    {
        private readonly MiContexto _context;

        // Constructor que recibe el contexto para interactuar con la base de datos
        public ControladorController(MiContexto context)
        {
            _context = context;
        }

        // Acción que maneja la vista de detalles del participante
        public async Task<IActionResult> Details(int id)
        {
            // Obtener el participante por su Id
            var participante = await _context.Participantes
                .Include(p => p.IdTareas) // Incluir las tareas asociadas al participante
                .ThenInclude(t => t.IdProyectoNavigation) // Incluir los proyectos asociados a las tareas
                .FirstOrDefaultAsync(p => p.IdParticipante == id);

            // Si el participante no existe, retornar un error
            if (participante == null)
            {
                return NotFound();
            }

            // Retornar la vista con el modelo de participante
            return View(participante);
        }

        // Acción para listar todos los participantes
        public async Task<IActionResult> Index()
        {
            // Obtener todos los participantes desde la base de datos
            var participantes = await _context.Participantes.ToListAsync();

            // Retornar la vista con la lista de participantes
            return View(participantes);
        }

        // Acción para crear un nuevo participante (si es necesario)
        [HttpGet]
        public IActionResult Create()
        {
            return View();
        }

        // Acción para manejar el envío de la creación de un nuevo participante
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("Nombre,Correo,Telefono")] Participante participante)
        {
            if (ModelState.IsValid)
            {
                // Agregar el nuevo participante a la base de datos
                _context.Add(participante);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index)); // Redirigir a la lista de participantes
            }

            return View(participante); // Si hay algún error, volver a mostrar el formulario
        }

        // Acción para editar un participante (si es necesario)
        [HttpGet]
        public async Task<IActionResult> Edit(int id)
        {
            // Buscar el participante por su Id
            var participante = await _context.Participantes.FindAsync(id);

            // Si no existe el participante, retornar error
            if (participante == null)
            {
                return NotFound();
            }

            return View(participante);
        }

        // Acción para manejar la edición de un participante
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("IdParticipante,Nombre,Correo,Telefono")] Participante participante)
        {
            if (id != participante.IdParticipante)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(participante); // Actualizar los datos en la base de datos
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!_context.Participantes.Any(p => p.IdParticipante == id))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index)); // Redirigir a la lista de participantes
            }

            return View(participante);
        }

        // Acción para eliminar un participante (si es necesario)
        [HttpGet]
        public async Task<IActionResult> Delete(int id)
        {
            var participante = await _context.Participantes
                .FirstOrDefaultAsync(p => p.IdParticipante == id);

            if (participante == null)
            {
                return NotFound();
            }

            return View(participante);
        }

        // Acción para manejar la eliminación de un participante
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var participante = await _context.Participantes.FindAsync(id);

            if (participante != null)
            {
                _context.Participantes.Remove(participante);
                await _context.SaveChangesAsync();
            }

            return RedirectToAction(nameof(Index)); // Redirigir a la lista de participantes
        }
    }
}
