/* =====================================================================
   CONCLUSIONES Y RECOMENDACIONES, Calculaser
   Se muestran al final del informe de la compañía (PDF) para el usuario
   cuyo bloque en users.js tiene:  conclusiones: 'calculaser'
   Para cambiar el texto, edita este archivo (no hay que tocar el código).
   Estructura de cada sección:
     titulo, intro (opcional), items: [ { cab, texto, lista: [...] } ]
     Un item puede tener solo "cab" (negrita), "texto" y/o "lista" (viñetas).
   ===================================================================== */
window.APP_CONCLUSIONES = window.APP_CONCLUSIONES || {};
window.APP_CONCLUSIONES.calculaser = {
  titulo: 'Conclusiones y recomendaciones',
  resumen: 'La primera evaluación de competencias deja una base muy valiosa para Calculaser. La organización tiene un nivel sólido de desempeño (3,22 sobre 4,0), una cultura de atención humanizada bien arraigada y un grupo directivo que su equipo valora. Las oportunidades se concentran en pensar y decidir con datos, en acompañar a algunos cargos clave y en afinar el proceso de medición.',
  secciones: [
    {
      titulo: 'Fortalezas', tipo: 'fortaleza',
      items: [
        { cab: 'Una organización con buen punto de partida.',
          texto: 'El 90% de las calificaciones está en los niveles 3 y 4, y solo el 9,6% se ubica en 1 o 2. Hay una base de competencias consolidada en las 104 personas evaluadas.' },
        { cab: 'Servicio humano y orientación al proceso.',
          texto: 'Atención humanizada (3,32) y Orientación a los procesos (3,28) son las competencias más sólidas con mayor volumen de datos. Esto es coherente con una institución de salud y es un activo que vale la pena cuidar.' },
        { cab: 'Liderazgo bien valorado por los equipos.',
          texto: 'Los líderes reciben en promedio 3,37 de sus equipos y 3,36 de sus jefes, por encima del resto de la organización (3,13 y 3,09). Se destacan:',
          lista: [
            'La Directora de gestión de calidad (4,00 de su equipo)',
            'La Directora DPS (3,78 de su jefe)',
            'El Director médico (3,68 de su equipo)',
            'El Regente de farmacia (3,78 de su jefe)',
            'El Supervisor del contact center (3,89 de su jefe)',
          ],
          cierre: 'Estos perfiles pueden ser referentes y mentores internos.' },
        { cab: 'Equipo con alto compromiso con el proceso.',
          texto: 'La participación fue amplia: 104 evaluados y 104 evaluadores, y los jefes completaron el 98% de sus calificaciones.' },
        { cab: 'Autopercepción prudente en varios líderes.',
          texto: 'La Directora contable y financiera y el Supervisor del contact center se autoevalúan por debajo de lo que perciben sus equipos y jefes. Es una muestra de humildad que se puede convertir en confianza con acompañamiento.' },
      ],
    },
    {
      titulo: 'Oportunidades de mejora', tipo: 'mejora', numerado: true,
      items: [
        { cab: 'Decidir y resolver con datos.',
          texto: 'Es la oportunidad más transversal. Según los jefes, en los cargos no directivos las calificaciones más bajas son Toma de decisiones (2,96), Autogestión (2,78) y Pensamiento analítico (3,00). Los comportamientos con menor puntaje son identificar causas, separar lo relevante y elegir la mejor opción con análisis.' },
        { cab: 'Resiliencia y liderazgo bajo presión.',
          texto: 'Está en 3,19 en directivos y 3,15 en el resto. Aparece como oportunidad sostener la claridad en escenarios de alta exigencia sin trasladar la tensión al equipo.' },
        { cab: 'Registros y documentación.',
          texto: 'Mantener los registros completos y confiables está entre los comportamientos más bajos. Mejorarlo fortalece la calidad y la seguridad del paciente.' },
        { cab: 'Acompañamiento a cargos específicos.',
          texto: 'Estos cargos se benefician de un apoyo más cercano:',
          lista: [
            'Programación de cirugía (líder 2,60; asistentes 2,90)',
            'Área contable y financiera (2,92)',
            'Dirección administrativa y comercial (3,14; el equipo la ve 0,65 por debajo de su autoevaluación)',
            'Cargos asistenciales de mayor volumen, por el número de personas',
          ],
          cierre: 'Con pocos evaluados por cargo, esto es una invitación a conversar y validar, no un juicio.' },
      ],
    },
    {
      titulo: 'Recomendaciones', tipo: 'neutro', salto: false,
      grupos: [
        { titulo: 'Corto plazo (0 a 3 meses): consolidar y conversar',
          lista: [
            'Entregar los resultados a cada área con enfoque de desarrollo, empezando por los directivos.',
            'Definir el nivel esperado por cargo para que los resultados se lean como brechas reales.',
            'Hacer una sesión de calibración con los jefes para unificar criterios de calificación.',
            'Ofrecer coaching individual a la líder de programación de cirugía y al Director administrativo y comercial, con un plan de 90 días.',
            'Dar un taller “Decidir con datos y con criterio” para los 12 líderes (2 sesiones de 4 horas, con un caso real de la clínica).',
            'Reforzar los registros con listas de verificación en programación de cirugía, farmacia, contabilidad y gestión de pacientes.',
            'Reconocer públicamente las fortalezas, especialmente la atención humanizada.',
          ] },
        { titulo: 'Mediano plazo (3 a 12 meses): desarrollar',
          lista: [
            'Programa de resiliencia y liderazgo bajo presión para directivos y líderes de procesos críticos (3 módulos de 3 horas).',
            'Formación en resolución de problemas con casos propios (5 porqués, espina de pescado) para auxiliares y profesionales, por grupos de 15 a 20 personas.',
            'Mentoría entre pares: que los referentes (Calidad, Dirección médica, DPS, Regencia de farmacia) acompañen a quienes están en proceso.',
            'Comunidad de práctica mensual de líderes para analizar casos reales.',
            'Entrenamiento en retroalimentación para jefes, para reducir la tendencia al 3 y los vacíos.',
            'Acompañamiento a la autoconfianza de la Directora contable y del Supervisor del contact center.',
            'Pulso trimestral en las competencias priorizadas y un instrumento más corto por familia de cargos.',
          ] },
        { titulo: 'Largo plazo (12 meses o más): sostener',
          lista: [
            'Segunda medición con la misma metodología, para ver la evolución.',
            'Integrar el modelo de competencias a selección, inducción, desempeño, carrera y sucesión.',
            'Rutas de carrera para cargos asistenciales y administrativos.',
            'Indicadores de talento en el tablero de gestión humana, junto con los de calidad y seguridad.',
          ] },
      ],
    },
  ],
};
