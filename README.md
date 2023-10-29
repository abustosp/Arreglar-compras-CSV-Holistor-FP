# Arreglo de compras CSV para el Holistor

Programa para arreglar las compras provenientes del CSV del Portai IVA compras, para emitir un archivo Excel (XLS) para que pueda ser utilizado para importar dicha información en el Holistor.

## Ejecución

Simplemente con descargar el .zip se puede ejecutar el programa (el link de descarga se encuentra a la derecha de GitHub, en la sección de "Releases")

---

### Funcionamiento

El programa va a pedir seleccionar la Carpeta donde se encuentren los CSV de Compras a procesar (lo ideal es tener en una carpeta todas las compras para procesarlos en un mismo lote)

El mismo se va a encargar de realizar varios pasos como:
- Eliminar Facturas tipo B
- Eliminar facturaas provenientes de una lista de cuits a eliminar (por ejemplo: se puede agregar el cuit de un supermercado debido a que existen clientes que suelen pedir facturas A en los mismos pero los bienes no se relacionan con su actividad comercial)
- Transformar a valores absolutos
- Adecuar la información para ser exportada
- Rellenar el XLS del modelo de importación de compras con la información procesada

---
### Desarrolladores


Desarrollado por: Federico Perret

Compilado por: Agustín Bustos Piasentini