/* =====================================================================
   USUARIOS DE LA APLICACIÓN
   ---------------------------------------------------------------------
   Cada usuario es una línea { ... }. Campos:
   - usuario: con el que se ingresa (no distingue mayúsculas)
   - clave:   contraseña (sí distingue mayúsculas)
   - nombre:  cómo se muestra en la barra superior
   - rol:     'admin' = carga los Excel a mano (pantalla "Cargar datos")
              'gh'    = Gestión Humana: entra directo al dashboard con
                        los datos ya cargados desde "archivos"
   - titulo:  título de la evaluación (barra superior, PDF y .zip)
   - campana: OPCIONAL. Valor exacto de EVALUATION_DESCRIPTION para que
              el usuario solo vea esa campaña. Vacío ('') = todo.
   - archivos: SOLO para rol 'gh'. Ruta, dentro del sitio publicado, de
              los dos Excel. Ver instrucciones en datos/LEEME.txt

   IMPORTANTE: este archivo lo puede leer cualquiera que conozca la
   dirección de la página. El inicio de sesión evita el acceso casual,
   pero no reemplaza una protección real del sitio.
   ===================================================================== */
window.APP_USERS = [
  { usuario: 'admin',         clave: 'Calculaser2026',  nombre: 'Administrador',  rol: 'admin',
    titulo: 'Evaluación de Competencias Calculaser', campana: '' },

  { usuario: 'talentohumano', clave: 'CambiaEstaClave', nombre: 'Talento Humano', rol: 'admin',
    titulo: 'Evaluación de Desempeño 2026', campana: '' },

  // ▼▼▼ USUARIO DE GESTIÓN HUMANA CALCULASER ▼▼▼
  // Cuando te entreguen los Excel, cópialos en la carpeta datos/calculaser/
  // con estos nombres exactos (o cambia aquí las rutas):
  { usuario: 'gh.calculaser',  clave: 'CambiaEstaClaveGH', nombre: 'Gestión Humana', rol: 'gh',
    titulo: 'Evaluación de Competencias Calculaser', campana: '',
    archivos: {
      calificaciones:    'datos/calculaser/calificaciones.xlsx',     // Excel con COMPETENCIA, CALIFICACION, AREA…
      retroalimentacion: 'datos/calculaser/retroalimentacion.xlsx',  // Excel con FEEDBACK y ASPECTS
    } },
  // ▲▲▲ FIN USUARIO GESTIÓN HUMANA ▲▲▲
];
