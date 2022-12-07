using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


Console.WriteLine("==-- SQUARE MATRIX --==\n" +
    "=- START -=\n" +
    "1) Create a matrix in the console. (Enter [1])\n" +
    "2) Load a matrix from a file. (Enter [2])");
string option = "";
while(option != "1" && option != "2")
{
    option = Console.ReadLine();
    switch (option)
    {
        case "1":
            ConsoleMatrix();
            break;
        case "2":
            FileMatrix();
            break;
    }
}


Console.ReadLine();

// used in file input specifically

static void FileMatrix()
{
    var filePath = GetAndCheckFile();
    bool matrixOk = false;

    while (matrixOk == false)
    {
        matrixOk = CheckMatrix(filePath);
        if (matrixOk == false)
        {
            filePath = GetAndCheckFile();
            matrixOk = CheckMatrix(filePath);
        }
    }

    string helpStr = WriteToHelpString(filePath);

    double sizeD = helpStr.Length;
    sizeD = Math.Sqrt(sizeD);
    int size = (int)sizeD;

    int[,] matrix = new int[size, size];

    matrix = WriteToMatrix(helpStr, size);

    Operations(matrix, size, filePath);

    Console.ReadLine();
}

static int[,] WriteToMatrix(string helpStr, int matrixSize)
{
    int[,] matrix = new int[matrixSize, matrixSize];
    int z = 0;
    for (int y = 0; y < matrixSize; y++)
    {
        for (int i = 0; i < matrixSize; i++)
        {
            int num;
            Int32.TryParse(helpStr[z].ToString(), out num);
            matrix[i, y] = num;
            z++;
        }
    }
    return matrix;
}

static string WriteToHelpString(string filePath)
{
    string[] lines = File.ReadAllLines(filePath);
    string helpString = "";
    foreach (string line in lines)
    {
        string[] col = line.Split(',');
        for (int i = 0; i < col.Length; i++)
        {
            helpString += col[i];
        }
    }
    return helpString;
}

static bool CheckMatrix(string filePath)
{
    string[] lines = File.ReadAllLines(filePath);
    int curNum = 0;
    foreach (string line in lines)
    {
        string[] col = line.Split(',');
        for (int i = 0; i < col.Length; i++)
        {
            if (col[i] == "")
            {
                Console.WriteLine($"Number {i+1} of the matrix is null!\n" +
                    "Invalid matrix.");
                return false;
            }
            else if (!int.TryParse(col[i], out curNum))
            {
                Console.WriteLine($"Number {i} of the matrix is invalid.\n" +
                    $"{col[i]} couldn't be converted to int.\n" +
                    "Invalid matrix.");
                return false;
            }
            else
            {
                continue;
            }
        }
    }
    return true;
}

static string GetAndCheckFile()
{
    Console.WriteLine("Please, enter the path for the file:");
    var filePath = Console.ReadLine();

    filePath = CheckIfFileExists(filePath);
    filePath = CheckFileExtension(filePath);

    string[] array = System.IO.File.ReadAllLines(filePath);
    while (array.Length == 0 || array.Length == 1)
    {
        if(array.Length == 0)
        {
            Console.WriteLine($"\nThe file is empty.\n" +
            "Please enter another file path.\n");
        }
        if(array.Length == 1)
        {
            Console.WriteLine($"\nThe file contains a single line.\n" +
            "A square matrix requires at least 2 lines.\n" +
            "Please enter another file path.\n");
        }
        filePath = Console.ReadLine();
        CheckIfFileExists(filePath);
        array = File.ReadAllLines(filePath);
    }
    return filePath;
}

// used in file input specifically



// used in console input specifically

static void ConsoleMatrix()
{
    int size = getSizeOfMatrix();
    int[,] matrix = new int[size, size];

    matrix = EnterValues(matrix, size);

    Operations(matrix, size, "0");
}

static int[,] EnterValues(int[,] matrix, int matrixSize)
{
    Console.WriteLine("\nPlease enter all the numbers that will be in the square matrix.\n" +
    "(Each column from up to down.)");

    for (int i = 0; i < matrixSize; i++)
    {
        Console.WriteLine($"\nColumn nr {i + 1}:");
        for (int y = 0; y < matrixSize; y++)
        {
            Console.WriteLine($"Row nr {y + 1}: ");
            matrix[i, y] = CheckNumberAndParseInt(Console.ReadLine());
        }
    }
    return matrix;
}

static int CheckNumberAndParseInt(string number)
{
    bool cont = false;
    int numberInt = 0;
    while(cont == false)
    {
        if (number == "")
        {
            Console.WriteLine("The input value is null. Please, try again.");
            number = Console.ReadLine();
        }
        else if (!int.TryParse(number, out numberInt))
        {
            Console.WriteLine("The input value couldn't be converted to int. Please, try again.");
            number = Console.ReadLine();
        }
        else
        {
            cont = true;
        }
    }
    return numberInt;
}

static int getSizeOfMatrix()
{
    bool cont = false;
    string sizeStr;
    int size = 0;
    Console.WriteLine("\nPlease enter the size of the matrix (from 2 to 10):");
    while (cont == false)
    {
        sizeStr = Console.ReadLine();
        if (sizeStr == "")
        {
            Console.WriteLine("\nThe input value is null.\nPlease, try again.");
        }
        else if (!int.TryParse(sizeStr, out size))
        {
            Console.WriteLine("\nThe input value couldn't be converted to int.\nPlease, try again.");
        }
        else if (size > 10)
        {
            Console.WriteLine("\nThe input value is greater than 10.\nPlease, try again.");
        }
        else if (size < 2)
        {
            Console.WriteLine("\nThe input value is lesser than 2.\nPlease, try again.");
        }
        else
        {
            cont = true;
        }
    }
    return size;
}

// used in console input specifically



// -- general use --

static void Operations(int[,] matrix, int matrixSize, string oldFilePath)
{
    Console.WriteLine("\n=- OPERATIONS -=\n" +
    ") Exit app. (Enter [/e])\n" +
    ") Write out the matrix. (Enter [/w])\n" +
    ") Sum elements of the main diagonal. (Enter [/sd])\n" +
    ") Sum elements of a given row. (Enter [/sr])\n" +
    ") Sum elements of a given column. (Enter [/sc])\n" +
    ") Write matrix to file. (Enter [/wf])\n");

    string option = "";
    while(option != "exit")
    {
        option = Console.ReadLine();
        switch (option)
        {
            case "/w":
                Console.WriteLine("--------------------");
                WriteOutMatrix(matrix, matrixSize);
                Console.WriteLine("--------------------");
                break;
            case "/sd":
                Console.WriteLine("--------------------");
                Console.WriteLine($"\nSum of the main dianogal: {SumMainDiagonal(matrix, matrixSize)}\n");
                Console.WriteLine("--------------------");
                break;
            case "/sr":
                Console.WriteLine("--------------------");
                Console.WriteLine($"\nSum of the row: {SumRow(matrix, matrixSize)}\n");
                Console.WriteLine("--------------------");
                break;
            case "/sc":
                Console.WriteLine("--------------------");
                Console.WriteLine($"\nSum of the column: {SumCol(matrix, matrixSize)}\n");
                Console.WriteLine("--------------------");
                break;
            case "/wf":
                Console.WriteLine("--------------------");
                string result = PrepareMatrixForFile(matrix, matrixSize);
                SaveResult(result, oldFilePath);
                Console.WriteLine("\n--------------------");
                break;
            case "/e":
                Environment.Exit(0);
                break;
        }
    }
}

static void WriteOutMatrix(int[,] matrix, int matrixSize)
{
    for (int y = 0; y < matrixSize; y++)
    {
        Console.Write("\n| ");
        for (int i = 0; i < matrixSize; i++)
        {
            Console.Write($"{matrix[i, y]} ");
        }
        Console.Write("|\n");
    }
    Console.Write("\n");
}


// mathematic operations

static int SumCol(int[,] matrix, int matrixSize)
{
    int col = 0;
    bool cont = false;

    Console.WriteLine("\nWhich column do you want to sum?");
    col = CheckNumberAndParseInt(Console.ReadLine());
    while (col <= 0 || col > matrixSize)
    {
        Console.WriteLine($"There is no such column as {col}.\n" +
                $"Possible columns are between 1 and {matrixSize}.\n" +
                $"Please, try again.\n");

        Console.WriteLine("\nWhich row do you want to sum?");
        col = CheckNumberAndParseInt(Console.ReadLine());
    }
    col--;
    int sum = 0;
    for (int i = 0; i < matrixSize; i++)
    {
        sum += matrix[col, i];
    }
    return sum;
}

static int SumRow(int[,] matrix, int matrixSize)
{
    int row = 0;
    bool cont = false;

    Console.WriteLine("\nWhich row do you want to sum?");
    row = CheckNumberAndParseInt(Console.ReadLine());
    while (row <= 0 || row > matrixSize)
    {
        Console.WriteLine($"\nThere is no such row as {row}.\n" +
                $"Possible rows are between 1 and {matrixSize}.\n" +
                $"Please, try again.\n");

        Console.WriteLine("\nWhich row do you want to sum?");
        row = CheckNumberAndParseInt(Console.ReadLine());
    }
    row--;
    int sum = 0;
    for (int i = 0; i < matrixSize; i++)
    {
        sum += matrix[i, row];
    }
    return sum;
}

static int SumMainDiagonal(int[,] matrix, int matrixSize)
{
    int sum = 0;
    for (int i = 0; i < matrixSize; i++)
    {
        sum += matrix[i, i];
    }
    return sum;
}

// mathematic operations


// file

static void SaveResult(string result, string oldFilePath)
{
    Console.WriteLine("Please, enter the path for the file:");
    var filePath = Console.ReadLine();

    while(filePath == oldFilePath)
    {
        Console.WriteLine("\nThe file you chose is the same file you read the matrix from.\n" +
            "There is no point in writing the same data to the file.\n" +
            "Please enter another file path.");
        filePath = Console.ReadLine();
    }

    filePath = CheckIfFileExists(filePath);

    WriteToFile(filePath, result);
}

static async void WriteToFile(string filePath, string result)
{
    if (new FileInfo(filePath).Length == 0)
    {
        await File.WriteAllTextAsync(filePath, result + "\n");
        Console.WriteLine("Result was written to the file successfully!\n");
    }
    else
    {
        Console.WriteLine("\nIt appears the file isn't empty!\n" +
            "Content of the file:\n");
        Console.WriteLine(File.ReadAllText(filePath));
        Console.WriteLine("\nWould you like to override this file?\n" +
            "(Yes - enter [1] No - enter [2]):");
        var option2 = Console.ReadLine();
        if (option2 == "1")
        {
            await File.WriteAllTextAsync(filePath, result + "\n");
            Console.WriteLine("Result was written to the file successfully!");
        }
        else if (option2 == "2")
        {
            Console.WriteLine("Please, enter the path for another file:");
            filePath = Console.ReadLine();

            filePath = CheckIfFileExists(filePath);

            WriteToFile(filePath, result);
        }
    }
}

static string CheckIfFileExists(string filePath)
{
    while (!File.Exists(filePath))
    {
        Console.WriteLine("File not found!\n" +
            "Please, try to enter the path again:");
        filePath = Console.ReadLine();
    }
    return filePath;
}

static string PrepareMatrixForFile(int[,] matrix, int matrixSize)
{
    string result = "";
    for(int i = 0; i < matrixSize; i++)
    {
        for(int y = 0; y < matrixSize-1; y++)
        {
            result += $"{matrix[y, i]},";
        }
        result += $"{matrix[matrixSize-1, i]}";
        result += Environment.NewLine;
    }
    result = result.Remove(result.LastIndexOf(Environment.NewLine));
    return result;
}

static string CheckFileExtension(string filePath)
{
    string ext = Path.GetExtension(filePath);
    while (ext != ".txt" && ext != ".csv")
    {
        Console.WriteLine($"\nThe file extension {ext} is not valid.\n" +
            "Please enter another file path.\n");
        filePath = Console.ReadLine();
        ext = Path.GetExtension(filePath);
    }
    return filePath;
}

// file


// -- general use --


// other, not used

// -- reading an .xlsx file
// (I don't have Office)
/* error
 *Could not load file or assembly 'office, Version=15.0.0.0, Culture=neutral, 
 *PublicKeyToken=71e9bc111e9429c' or one of its dependencies. 
 *The system cannot find the file specified
*/
/* the function
static void ReadXLSXFile(string filePath)
{
    CheckIfFileExists(filePath);
    //Create COM Objects. Create a COM object for everything that is referenced
    Excel.Application xlApp = new Excel.Application();
    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
    Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
    Excel.Range xlRange = xlWorksheet.UsedRange;

    int rowCount = xlRange.Rows.Count;
    int colCount = xlRange.Columns.Count;

    for (int i = 1; i <= rowCount; i++)
    {
        for (int j = 1; j <= colCount; j++)
        {
            //new line
            if (j == 1)
                Console.Write("\r\n");

            //write the value to the console
            if (xlRange.Cells[i, j] != null)
                Console.Write(xlRange.Cells[i, j].ToString() + "\t");
        }
    }
    GC.Collect();
    GC.WaitForPendingFinalizers();

    Marshal.ReleaseComObject(xlRange);
    Marshal.ReleaseComObject(xlWorksheet);

    xlWorkbook.Close();
    Marshal.ReleaseComObject(xlWorkbook);

    xlApp.Quit();
    Marshal.ReleaseComObject(xlApp);
}
*/