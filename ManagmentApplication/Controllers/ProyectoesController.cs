using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ManagmentApplication.Data;
using ManagmentApplication.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Net.Http;
using System.IO;

namespace ManagmentApplication.Controllers
{
    public class ProyectoesController : Controller
    {
        private readonly MiContexto _context;

        public ProyectoesController(MiContexto context)
        {
            _context = context;
        }

        // GET: Proyectoes
        public async Task<IActionResult> Index()
        {
            return View(await _context.Proyectos.ToListAsync());
        }

        // Acción para exportar a Excel
        public async Task<IActionResult> ExportToExcel()
        {
            var proyectos = await _context.Proyectos.ToListAsync();

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Proyectos");

                // Encabezado
                worksheet.Cells[1, 1].Value = "Id Proyecto";
                worksheet.Cells[1, 2].Value = "Nombre";
                worksheet.Cells[1, 3].Value = "Descripción";
                worksheet.Cells[1, 4].Value = "Fecha de Creación";
                worksheet.Cells[1, 5].Value = "Fecha Estimada de Fin";
                worksheet.Cells[1, 6].Value = "Estado";
                worksheet.Cells[1, 7].Value = "Imagen";

                // Estilo del encabezado
                using (var range = worksheet.Cells[1, 1, 1, 7])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                // Llenar datos y agregar imágenes
                for (int i = 0; i < proyectos.Count; i++)
                {
                    var proyecto = proyectos[i];
                    worksheet.Cells[i + 2, 1].Value = proyecto.IdProyecto;
                    worksheet.Cells[i + 2, 2].Value = proyecto.Nombre;
                    worksheet.Cells[i + 2, 3].Value = proyecto.Descripcion;
                    worksheet.Cells[i + 2, 4].Value = proyecto.FechaCreacion?.ToString("yyyy-MM-dd");
                    worksheet.Cells[i + 2, 5].Value = proyecto.FechaFinEstimada?.ToString("yyyy-MM-dd");
                    worksheet.Cells[i + 2, 6].Value = proyecto.Estado;

                    // Ajustar la altura de la fila para que las imágenes puedan caber correctamente
                    worksheet.Row(i + 2).Height = 100;  // Puedes ajustar la altura según tus necesidades

                    // Agregar imagen si existe la URL
                    if (!string.IsNullOrEmpty(proyecto.ImagenUrl))
                    {
                        try
                        {
                            // Descargar la imagen desde la URL
                            using (var httpClient = new HttpClient())
                            {
                                var imageBytes = await httpClient.GetByteArrayAsync(proyecto.ImagenUrl);

                                // Agregar la imagen al archivo Excel usando el flujo de bytes
                                var excelImage = worksheet.Drawings.AddPicture($"Image{i + 1}", new MemoryStream(imageBytes));

                                // Posicionar la imagen en la fila y la columna de la imagen
                                excelImage.SetPosition(i + 1, 0, 6, 0);  // Fila i+1 y columna 7 (columna "Imagen")

                                // Ajustar el tamaño de la imagen al tamaño de la celda
                                excelImage.SetSize(80, 80);  // Ajustar el tamaño de la imagen a 80x80 píxeles, puedes modificar estos valores

                            }
                        }
                        catch (Exception ex)
                        {
                            // Si hay un error con la URL de la imagen, manejarlo aquí
                            worksheet.Cells[i + 2, 7].Value = "Imagen no válida";
                            Console.WriteLine($"Error al cargar la imagen para el proyecto {proyecto.Nombre}: {ex.Message}");
                        }
                    }
                    else
                    {
                        worksheet.Cells[i + 2, 7].Value = "No hay imagen";
                    }
                }

                // Ajustar el tamaño de las columnas
                worksheet.Cells.AutoFitColumns();

                // Convertir el archivo a un arreglo de bytes y devolverlo
                var fileContents = package.GetAsByteArray();
                return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Proyectos.xlsx");
            }
        }



        // GET: Proyectoes/Details/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var proyecto = await _context.Proyectos
                .FirstOrDefaultAsync(m => m.IdProyecto == id);
            if (proyecto == null)
            {
                return NotFound();
            }

            return View(proyecto);
        }

        // GET: Proyectoes/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: Proyectoes/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("IdProyecto,Nombre,Descripcion,FechaCreacion,FechaFinEstimada,Estado,ImagenUrl")] Proyecto proyecto)
        {
            if (ModelState.IsValid)
            {
                _context.Add(proyecto);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            return View(proyecto);
        }

        // GET: Proyectoes/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var proyecto = await _context.Proyectos.FindAsync(id);
            if (proyecto == null)
            {
                return NotFound();
            }
            return View(proyecto);
        }

        // POST: Proyectoes/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("IdProyecto,Nombre,Descripcion,FechaCreacion,FechaFinEstimada,Estado,ImagenUrl")] Proyecto proyecto)
        {
            if (id != proyecto.IdProyecto)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(proyecto);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!ProyectoExists(proyecto.IdProyecto))
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
            return View(proyecto);
        }

        // GET: Proyectoes/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var proyecto = await _context.Proyectos
                .FirstOrDefaultAsync(m => m.IdProyecto == id);
            if (proyecto == null)
            {
                return NotFound();
            }

            return View(proyecto);
        }

        // POST: Proyectoes/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var proyecto = await _context.Proyectos.FindAsync(id);
            if (proyecto != null)
            {
                _context.Proyectos.Remove(proyecto);
            }

            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool ProyectoExists(int id)
        {
            return _context.Proyectos.Any(e => e.IdProyecto == id);
        }
    }
}
