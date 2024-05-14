#pragma warning disable 0219 // Desactiva la advertencia del compilador para variables declaradas pero no utilizadas.


using System; // Proporciona tipos fundamentales y clases base comunes en .NET.
using System.Collections.Generic; // Incluye definiciones para colecciones genéricas como listas y diccionarios.
using System.IO; // Contiene clases para el manejo de archivos y flujos de datos.
using System.Runtime.CompilerServices; // Ofrece funcionalidades para el control avanzado del comportamiento del compilador.
using System.Text; // Incluye clases para la manipulación avanzada de textos y codificaciones.
using NPOI.SS.UserModel; // Proporciona interfaces comunes para manejar archivos Excel.
using NPOI.HSSF.UserModel; // Permite trabajar con archivos Excel en formato .xls.
using NPOI.XSSF.UserModel; // Permite trabajar con archivos Excel en formato .xlsx.
using UnityEditor; // Contiene herramientas para desarrollar y extender el editor de Unity.
using UnityEngine; // Incluye las clases esenciales para el desarrollo de juegos en Unity.
using System.Linq; // Proporciona extensiones para operaciones de consulta y manipulación de colecciones.


// Clase principal que lee y procesa archivos Excel
public class ReaderExcel : MonoBehaviour
{
    // Clase para almacenar información detallada sobre un archivo Excel.
    [System.Serializable]
    private class ExcelFile
    {
        public string fileName; // Nombre del archivo.
        public string filePath; // Ruta completa del archivo.
        public string fileRelativePath; // Ruta relativa del archivo.
        public string fileExtension; // Extensión del archivo.
    }

    // Información básica de una hoja de cálculo
    private class SheetInfo
    {
        public string name; // Nombre de la hoja.
    }

    // Información sobre los datos dentro de cada celda
    private class DataInfo
    {
        public string desc; // Descripción del dato.
        public bool isArray; // Indica si el dato es un arreglo.
        public bool isEnable; // Indica si el dato está habilitado.
        public string titleName; // Nombre del título de la columna.
        public ValueType type; // Tipo de dato almacenado.
        public int row; // Número de fila donde se encuentra el dato.
        public string value; // Valor leído de la celda.
    }

    // Enumeración para tipos de datos posibles en las celdas
    private enum ValueType
    {
        BOOL,
        STRING,
        INT,
        FLOAT,
        DOUBLE
    }

    private readonly Dictionary<string, List<DataInfo>> dataTitleDic = new Dictionary<string, List<DataInfo>>(); // Diccionario para almacenar los títulos y sus respectivos datos leídos.
    private static readonly ExcelFile excelFile = new ExcelFile(); // Archivo Excel actual
    private readonly List<SheetInfo> sheetList = new List<SheetInfo>(); // Lista de hojas en el archivo

    // Editor personalizado para la clase ReaderExcel en Unity.
#if UNITY_EDITOR
    [CustomEditor(typeof(ReaderExcel))]
    public class ReaderExcelEditor : Editor
    {
        // Interfaz de usuario en el Inspector para seleccionar el archivo Excel
        public override void OnInspectorGUI()
        {
            base.OnInspectorGUI();

            ReaderExcel fileSelector = (ReaderExcel)target;

            EditorGUILayout.Space();

            if (GUILayout.Button("Select Excel File"))
            {
                fileSelector.SelectExcelFile();
            }
        }
    }
#endif

    private void Awake()
    {
        // Recuperar la ruta del archivo seleccionado de las preferencias del editor
        string selectedFilePath = EditorPrefs.GetString("SelectedExcelFilePath", "");

        if (!string.IsNullOrEmpty(selectedFilePath))
        {
            // Ruta y detalles del archivo
            excelFile.filePath = selectedFilePath;
            string relativePath = "Assets" + excelFile.filePath.Substring(Application.dataPath.Length);
            excelFile.fileRelativePath = relativePath;
            excelFile.fileName = Path.GetFileName(excelFile.fileRelativePath);
            excelFile.fileExtension = Path.GetExtension(excelFile.fileName);

            Debug.Log("Ruta recuperada: " + excelFile.fileRelativePath);
        }
    }

    void Start() // Inicialización de la lectura del archivo Excel.
    {
        if (excelFile.fileRelativePath != "")
        {
            Debug.Log("Verificando Ruta 2: " + excelFile.fileRelativePath);

            // Leer y procesar el archivo Excel
            ReadExcelMethod();
            AssetDatabase.ImportAsset(excelFile.fileRelativePath);
            AssetDatabase.Refresh();
        }
        else
        {
            Debug.LogError("No se seleccionó ningún archivo Excel.");
        }
    }

    public void SelectExcelFile() // Método para seleccionar un archivo Excel.
    {
        // Selector de archivos para elegir un archivo Excel
        string selectedFilePath = EditorUtility.OpenFilePanel("Select Excel File", "", "xls,xlsx");

        if (!string.IsNullOrEmpty(selectedFilePath))
        {
            if (Path.GetExtension(selectedFilePath).Equals(".xls") || Path.GetExtension(selectedFilePath).Equals(".xlsx"))
            {
                // Procesar y guardar ruta del archivo seleccionado
                excelFile.filePath = selectedFilePath;
                string relativePath = "Assets" + excelFile.filePath.Substring(Application.dataPath.Length);
                excelFile.fileRelativePath = relativePath;
                excelFile.fileName = Path.GetFileName(excelFile.fileRelativePath);
                excelFile.fileExtension = Path.GetExtension(excelFile.fileName);
                EditorPrefs.SetString("SelectedExcelFilePath", excelFile.filePath);

                Debug.Log("Ruta: " + excelFile.fileRelativePath);
            }
            else
            {
                Debug.LogWarning("Formato invalido. Por favor selecciona un archivo .xls o .xlsx");
            }
        }
    }

    private void ReadExcelMethod() // Método para leer el contenido de un archivo Excel.
    {
        Debug.Log("Verificando Ruta 3: " + excelFile.fileRelativePath);
        using (var stream = File.Open(excelFile.fileRelativePath, FileMode.Open, FileAccess.Read))
        {
            IWorkbook book = null;

            // Seleccionar el tipo de libro de trabajo en función de la extensión del archivo
            if (Path.GetExtension(excelFile.fileExtension) == ".xls")
            {
                book = new HSSFWorkbook(stream);
            }
            else if (Path.GetExtension(excelFile.fileExtension) == ".xlsx")
            {
                book = new XSSFWorkbook(stream);
            }
            else
            {
                Debug.LogError("Formato de archivo Excel no compatible.");
                return;
            }

            // Procesar cada hoja en el libro de trabajo
            for (int j = 0; j < book.NumberOfSheets; j++)
            {
                var sheet = book.GetSheetAt(j);
                SheetInfo sheetInfo = new SheetInfo { name = sheet.SheetName };
                sheetList.Add(sheetInfo);

                var title = sheet.GetRow(0);
                if (title == null)
                {
                    Debug.LogWarning("Archivo sin nombre de las variables a utilizar.");
                    return;
                }

                // Diccionario para el conteo de tipos de datos por columna
                Dictionary<int, Dictionary<ValueType, int>> columnTypeCounts = new Dictionary<int, Dictionary<ValueType, int>>();
                List<DataInfo> dataList = new List<DataInfo>();

                // Leer y analizar datos de cada fila
                for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    var dataRow = sheet.GetRow(rowIndex);
                    if (dataRow == null)
                    {
                        Debug.LogWarning($"La fila {rowIndex} no contiene datos.");
                        continue;
                    }

                    for (int i = 0; i < title.Cells.Count; i++)
                    {
                        if (!string.IsNullOrEmpty(title.Cells[i].ToString()))
                        {
                            var data = new DataInfo
                            {
                                titleName = title.Cells[i].ToString(),
                                isArray = title.Cells[i].ToString().Contains("[]")
                            };
                            if (data.isArray)
                            {
                                data.titleName = data.titleName.Replace("[]", "");
                                Debug.Log($"Se detectó un array en la variable: {data.titleName} en la fila de títulos");
                            }

                            var cell = dataRow.Cells[i];
                            data = SetDataType(cell, data);
                            data.row = rowIndex + 1;

                            if (!columnTypeCounts.ContainsKey(i))
                            {
                                columnTypeCounts[i] = new Dictionary<ValueType, int>();
                            }

                            if (!columnTypeCounts[i].ContainsKey(data.type))
                            {
                                columnTypeCounts[i][data.type] = 0;
                            }
                            columnTypeCounts[i][data.type]++;

                            dataList.Add(data);
                        }
                    }
                }

                dataTitleDic[sheetInfo.name] = dataList;
                PrintDataFromDictionary();

                // Evaluar y validar los tipos de datos más frecuentes en cada columna
                foreach (var colIndex in columnTypeCounts.Keys)
                {
                    var mostFrequentType = columnTypeCounts[colIndex].OrderByDescending(kvp => kvp.Value).First().Key;
                    Debug.Log($"Columna {colIndex + 1} en hoja '{sheetInfo.name}' tiene tipo mayoritario '{mostFrequentType}'.");

                    foreach (var dataInfo in dataList.Where(d => d.titleName == title.Cells[colIndex].ToString()))
                    {
                        if (dataInfo.type != mostFrequentType)
                        {
                            Debug.LogWarning($"Inconsistencia de tipo detectada en hoja '{sheetInfo.name}', columna {colIndex + 1}, fila {dataInfo.row} (esperado: {mostFrequentType}, encontrado: {dataInfo.type}).");
                        }
                    }
                }
            }
        }
    }

    // Establecer el tipo de dato en función del contenido de la celda
    private static DataInfo SetDataType(ICell cell, DataInfo data)
    {
        data.value = cell.ToString();
        data.isArray = data.titleName.Contains("[]");
        if (data.isArray)
        {
            data.titleName = data.titleName.Replace("[]", "");
        }

        if (data.isArray && cell.ToString().Contains(","))
        {
            string[] arrayValues = cell.ToString().Split(',');
            foreach (string value in arrayValues)
            {
                if (int.TryParse(value, out _))
                {
                    data.type = ValueType.INT;
                }
                else if (float.TryParse(value, out _))
                {
                    data.type = ValueType.FLOAT;
                }
                else if (double.TryParse(value, out _))
                {
                    data.type = ValueType.DOUBLE;
                }
                else if (bool.TryParse(value, out _))
                {
                    data.type = ValueType.BOOL;
                }
                else
                {
                    data.type = ValueType.STRING;
                    break;
                }
            }
        }
        else
        {
            if (int.TryParse(cell.ToString(), out _))
            {
                data.type = ValueType.INT;
            }
            else if (float.TryParse(cell.ToString(), out _))
            {
                data.type = ValueType.FLOAT;
            }
            else if (double.TryParse(cell.ToString(), out _))
            {
                data.type = ValueType.DOUBLE;
            }
            else if (bool.TryParse(cell.ToString(), out _))
            {
                data.type = ValueType.BOOL;
            }
            else
            {
                data.type = ValueType.STRING;
            }
        }

        return data;
    }

    // Imprimir datos leídos para depuración
    private void PrintDataFromDictionary()
    {
        Debug.Log("Imprimiendo datos almacenados:");
        foreach (var sheet in dataTitleDic.Keys)
        {
            Debug.Log($"Hoja: {sheet}");
            foreach (var dataInfo in dataTitleDic[sheet])
            {
                Debug.Log($"Fila: {dataInfo.row}, Columna: '{dataInfo.titleName}, Tipo: {dataInfo.type}, Valor: '{dataInfo.value}'");
            }
        }
    }
}
