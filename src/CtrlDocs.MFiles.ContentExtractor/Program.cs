using MFilesAPI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CtrlDocs.MFiles.ContentExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            PrintInit();
            PrintHelp();
            if (args.Length < 1)
            {
                Environment.Exit(0);
            }

            // Class folder check
            var classFolders = Directory
                .GetDirectories(args[0])
                .Select((d, i) => new { d, i });
            Console.WriteLine($"Found these class folders, using the path supplied ({args[0]}):");
            foreach (var folder in classFolders)
            {
                Console.WriteLine($"{folder.i + 1}.\tName: {folder.d.Split('\\').Last()}, Folder: {folder.d}");
            }

            Console.WriteLine("\nPlease ensure that these are correct before continuing...");
            EnterToContinue();

            // Client setup
            var client = new MFilesClientApplication();
            var vaultConnections = client
                .GetVaultConnections()
                .Cast<VaultConnection>()
                .Where(c => c.IsLoggedIn())
                .Select((c, i) => new { c, i });

            // Vault listing
            Console.WriteLine("Found these connected and logged in vaults:");
            foreach (var connection in vaultConnections)
            {
                Console.WriteLine($"{connection.i + 1}.\tGUID: {connection.c.GetGUID()}, Name: {connection.c.Name}");
            }

            // Vault selection
            Console.WriteLine($"\n>Please select a number (1 through {vaultConnections.Count()}):");
            var vaultNumber = int.Parse(Console.ReadLine());
            var vaultConnection = vaultConnections.First(c => c.i == vaultNumber - 1).c;
            Console.WriteLine($"You selected vault GUID: {vaultConnection.GetGUID()}, Name: {vaultConnection.Name}");

            // Vault binding
            Console.Write("Binding to the vault...");
            var vault = vaultConnection.BindToVault(IntPtr.Zero, true, true);
            Console.WriteLine(" Success!");

            // loop
            while (true)
            {
                Console.WriteLine("\n===================================================");

                // Document selection
                Console.WriteLine($">Please select a Document object ID to convert to text:");
                var documentNumber = int.Parse(Console.ReadLine());
                var objID = new ObjID() { Type = (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument, ID = documentNumber };
                var objVer = vault.ObjectOperations.GetLatestObjVer(objID, true);
                var file = vault.ObjectFileOperations.GetFiles(objVer)[1];
                Console.WriteLine($"You selected file ID: {documentNumber}, Name: {file.Title}");

                // Content retrieval
                Console.Write("Retrieving file content...");
                var contents = vault.ObjectFileOperations.GetTextContentForFile(objVer, file.FileVer);
                Console.WriteLine(" Success!");

                // Class selection
                Console.WriteLine($"\n>Please select a class for this content (1 through {classFolders.Count()}):");
                foreach (var folder in classFolders)
                {
                    Console.WriteLine($"{folder.i + 1}.\tName: {folder.d.Split('\\').Last()}");
                }
                var classNumber = int.Parse(Console.ReadLine());
                var classFolder = classFolders.First(cf => cf.i == classNumber - 1).d;
                Console.WriteLine($"You selected class Name: {classFolder.Split('\\').Last()}");

                // Content Save
                var directoryPath = $@"{classFolder}\RequiresCensoring";
                Directory.CreateDirectory(directoryPath);
                var filepath = $@"{directoryPath}\{file.GetNameForFileSystem()}.txt";
                File.WriteAllText(filepath, contents);
                Console.WriteLine($"Contents successfully saved to: {filepath}");
            }

            EnterToContinue();
        }

        static void PrintInit()
        {
            Console.WriteLine(
                "===================================================\n" +
                "= Content Extractor for M-Files v1.0\n" +
                "==================================================="
                );
        }

        static void PrintHelp()
        {
            Console.WriteLine(
                "= Usage instructions:\n" + 
                "= 1. Provide the script with a path to the 'Classes' folder, relative to the location of this script.\n" + 
                "= 2. Ensure that the target vault is logged in and ready to use from the M-Files client.\n" + 
                "= 3. Once started, start by selecting a vault, enter a document object ID and lastly the class that it belongs to.\n" +
                "= 4. Files will be placed in a subfolder called 'RequiresCensoring' where they should stay until properly censored.\n" +
                "= 5. For proper document censoring, refer to the instructions on GitHub.\n" + 
                "==================================================="
                );
            EnterToContinue();
        }

        static void EnterToContinue()
        {
            Console.WriteLine("Press enter to continue...");
            Console.ReadLine();
        }
    }
}
