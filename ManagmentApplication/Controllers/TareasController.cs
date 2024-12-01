using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using ManagmentApplication.Data;
using ManagmentApplication.Models;

namespace ManagmentApplication.Controllers
{
    public class TareasController : Controller
    {
        private readonly MiContexto _context;

        public TareasController(MiContexto context)
        {
            _context = context;
        }

        // GET: Tareas/Create
        public IActionResult Create()
        {
            ViewData["IdProyecto"] = new SelectList(_context.Proyectos, "IdProyecto", "IdProyecto");
            return View();
        }

        // POST: Tareas/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("IdTarea,IdProyecto,Nombre,Descripcion,TiempoRealizado,TiempoEsperado,FechaCreacion,Estado")] Tarea tarea)
        {
            
            
                // Verifica que 'Estado' tenga un valor válido
                if (string.IsNullOrWhiteSpace(tarea.Estado))
                {
                    tarea.Estado = "Pendiente"; // Valor por defecto
                }

                // Aquí buscamos el Proyecto correspondiente
                var proyecto = await _context.Proyectos.FindAsync(tarea.IdProyecto);
                if (proyecto != null)
                {
                    // Asignamos el Proyecto a la propiedad de navegación
                    tarea.IdProyectoNavigation = proyecto;
                }
                else
                {
                    // Si no encontramos el Proyecto, devolvemos un error o mensaje adecuado
                    ModelState.AddModelError("", "El proyecto no existe.");
                    ViewData["IdProyecto"] = new SelectList(_context.Proyectos, "IdProyecto", "IdProyecto", tarea.IdProyecto);
                    return View(tarea);
                }

                // Agregar la tarea al contexto y guardar cambios
                _context.Add(tarea);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            

            // Si el modelo no es válido, mostramos la vista nuevamente con los datos
            ViewData["IdProyecto"] = new SelectList(_context.Proyectos, "IdProyecto", "Nombre", tarea.IdProyecto);
            return View(tarea);
        }


        // GET: Tareas/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var tarea = await _context.Tareas.FindAsync(id);
            if (tarea == null)
            {
                return NotFound();
            }

            ViewData["IdProyecto"] = new SelectList(_context.Proyectos, "IdProyecto", "IdProyecto", tarea.IdProyecto);
            return View(tarea);
        }

        // POST: Tareas/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("IdTarea,IdProyecto,Nombre,Descripcion,TiempoRealizado,TiempoEsperado,FechaCreacion,Estado")] Tarea tarea)
        {
            if (id != tarea.IdTarea)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    // Verifica que 'Estado' tenga un valor válido
                    if (string.IsNullOrWhiteSpace(tarea.Estado))
                    {
                        tarea.Estado = "Pendiente"; // Valor por defecto
                    }

                    _context.Update(tarea);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!TareaExists(tarea.IdTarea))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            ViewData["IdProyecto"] = new SelectList(_context.Proyectos, "IdProyecto", "IdProyecto", tarea.IdProyecto);
            return View(tarea);
        }

        // GET: Tareas
        public async Task<IActionResult> Index()
        {
            var tareas = _context.Tareas.Include(t => t.IdProyectoNavigation);
            return View(await tareas.ToListAsync());
        }

        private bool TareaExists(int id)
        {
            return _context.Tareas.Any(e => e.IdTarea == id);
        }
    }
}
