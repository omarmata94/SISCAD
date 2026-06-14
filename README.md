# SISCAD - Registro de Evaluación Continua

Aplicación en Python para cargar, editar y guardar registros de evaluación continua usando el formato de tu archivo `ListaDeCalificaciones3E.xlsx`.

## Requisitos

- Python 3.10 o superior
- `pip` instalado

## Instalación

```bash
cd /Users/omarmata/Documents/Desarrollos/SISCAD
python -m pip install -r requirements.txt
```

## Ejecución

```bash
streamlit run app.py
```

## Uso

1. Coloca tu archivo `ListaDeCalificaciones3E.xlsx` en el directorio del proyecto.
2. Abre la aplicación en el navegador.
3. Si el archivo existe localmente, pulsa el botón "Usar archivo local ListaDeCalificaciones3E.xlsx".
4. Edita los datos generales, la lista de alumnos y los trimestres.
5. Guarda el libro actualizado y descárgalo.5. Exporta el formato que secretaría solicita, pudiendo elegir el primer, segundo o los tres trimestres.
6. Si eliges los tres trimestres, la app también genera automáticamente la hoja `FINAL` con el resumen para secretaría.
## Qué incluye esta aplicación

- Lectura de la hoja `DATOS` con la lista de alumnos.
- Edición de los tres trimestres con detección dinámica de la estructura del archivo: actividades, tareas, proyectos, participación, conducta, examen y promedio final.
- Cálculo automático de promedios y resultados.
- Guardado en un nuevo archivo Excel `registro_evaluacion_actualizado.xlsx`.
