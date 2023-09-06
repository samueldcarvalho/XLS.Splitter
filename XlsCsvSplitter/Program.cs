using Microsoft.Office.Interop.Excel;

public class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine("Primeiramente digite o caminho do arquivo que gostaria splittar:");
        Console.WriteLine("*Exemplo: C:\\Planilhas\\Planilha.xls\n");

        Application excel;
        Workbook workbook;

        try
        {
            excel = new Application();
            excel.Visible = true;
        }
        catch (Exception)
        {
            Console.WriteLine("Excel não pode ser iniciado. Verifique se o Office está instalado.");
            throw;
        }

        while (true)
        {
            string? filePath = Console.ReadLine();
            Console.WriteLine("");

            if (string.IsNullOrWhiteSpace(filePath))
            {
                Console.WriteLine("O caminho é obrigatório. Tente novamente:\n");
                continue;
            }

            if (!File.Exists(filePath))
            {
                Console.WriteLine("Arquivo não encontrado. Tente novamente:\n");
                continue;
            }

            try
            {
                workbook = excel.Workbooks.Open(filePath);
            }
            catch
            {
                Console.WriteLine("Arquivo inválido, tente novamente:\n");
                continue;
            }

            break;
        }

        Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

        int rowsCount = worksheet.UsedRange.Rows.Count;

        Console.WriteLine($"{rowsCount} linhas encontradas.\n");

        Console.WriteLine("Digite um valor caso queira remover o cabeçalho.");
        Console.WriteLine("*Digite 0 para manter o cabeçalho.\n");

        int headerCount = 0;

        while (true)
        {
            if (!int.TryParse(Console.ReadLine(), out headerCount))
            {
                Console.WriteLine("");
                Console.WriteLine("Digite um número válido:\n");
                continue;
            }

            Console.WriteLine("");
            break;
        }

        Console.WriteLine("Agora vamos splittar o arquivo. \nPrimeiro, digite a quantidade máxima de registros por arquivo:\n");

        int rowsPerFile = 0;

        while (true)
        {
            if(!int.TryParse(Console.ReadLine(), out rowsPerFile))
            {
                Console.WriteLine("");
                Console.WriteLine("Digite um número válido:\n");
                continue;
            }

            Console.WriteLine("");
            break;
        }

        Console.WriteLine("Agora digite o caminho para salvar os novos arquivos:");
        Console.WriteLine("*Exemplo: C:\\Planilhas\\Splittadas\n");

        string? savePath = string.Empty;

        while (true)
        {
            savePath = Console.ReadLine();
            Console.WriteLine("");

            if (string.IsNullOrWhiteSpace(savePath))
            {
                Console.WriteLine("O caminho para salvar é obrigatório. Tente novamente:\n");
                continue;
            }

            if(!Path.Exists(savePath))
            {
                Console.WriteLine("Caminho não encontrado. Digite novamente:\n");
                continue;
            }

            break;
        }

        Console.Clear();
        Console.WriteLine("Calculando.");
        Thread.Sleep(1000);
        Console.Clear();
        Console.WriteLine("Calculando..");
        Thread.Sleep(1000);
        Console.Clear();
        Console.WriteLine("Calculando...");
        Thread.Sleep(1000);
        Console.Clear();

        int recordRows = rowsCount - headerCount;

        int filesCount = recordRows / rowsPerFile;
        int extraRecords = recordRows % rowsPerFile;

        int totalFilesCount = filesCount + (extraRecords > 0 ? 1 : 0);

        Console.WriteLine("O arquivo será dividido da seguinte forma:");

        Console.WriteLine(@$"
            Registros no arquivo original: {recordRows}
            Quantidade de novos arquivos: {totalFilesCount}
            Serão salvos em: {Path.GetFullPath(savePath)}

        Deseja confirmar?

        *Pressione ENTER para confirmar.
        *Pressione ESC para cancelar.");

        Console.WriteLine();
        
        while (true)
        {
            if (Console.ReadKey().Key == ConsoleKey.Enter)
            {
                Console.Clear();
                break;
            }

            if (Console.ReadKey().Key == ConsoleKey.Escape)
            {
                workbook.Close();
                excel.Quit();
                return;
            }
        }

        if(filesCount > 0)
        {
            for(int i = 0; i < filesCount; i++) 
            {
                Console.WriteLine($"Gerando arquivos: {(100 / totalFilesCount) * i}%");

                int linhaInicial = (i * rowsPerFile) + 1 + (i == 0 ? headerCount : 0);
                int linhaFinal = (i * rowsPerFile) + rowsPerFile;

                Workbook newWorkbook = excel.Workbooks.Add();
                Worksheet newWorksheet = (Worksheet)newWorkbook.Worksheets[1];

                var rows = worksheet.Range[$"A{linhaInicial}", $"A{linhaFinal}"].EntireRow;

                newWorksheet.Range["A1", $"A{rowsPerFile - 1}"].EntireRow.Value = rows.Value;

                newWorkbook.SaveAs2(Path.Combine(savePath, $"{linhaInicial}-{linhaFinal}.xls"));
                newWorkbook.Close();

                Console.Clear();
            }
        }

        if(extraRecords > 0)
        {
            Console.WriteLine($"Gerando arquivos: 98%");

            int linhaInicial = ((filesCount) * rowsPerFile) + 1;
            int linhaFinal = ((filesCount) * rowsPerFile) + extraRecords;

            Workbook newWorkbook = excel.Workbooks.Add();
            Worksheet newWorksheet = (Worksheet)newWorkbook.Worksheets[1];

            var rows = worksheet.Range[$"A{linhaInicial}", $"A{linhaFinal}"].EntireRow;

            newWorksheet.Range["A1", $"A{extraRecords}"].EntireRow.Value = rows.Value;

            newWorkbook.SaveAs2(Path.Combine(savePath, $"{linhaInicial}-{linhaFinal}.xls"));
            newWorkbook.Close();

            Console.Clear();
        }

        workbook.Close();
        excel.Quit();
    }
}