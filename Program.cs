using System.CommandLine;
using System.Drawing.Text;
using System.Runtime.Versioning;

[SupportedOSPlatform("windows")]
class Program
{
    static async Task<int> Main(string[] args)
    {
        var rootCommand = new RootCommand("Extracts Adobe Creative Cloud fonts to be used in other programs.");
        var copyCommand = new Command("copy", "Copy files from Creative Cloud directory.");

        var dryRunOption = new Option<bool>
            (name: "--dry",
            description: "Perform a dry run and don't copy files. Useful for troubleshooting.");
        dryRunOption.AddAlias("-d");
        var verboseOption = new Option<bool>
            (name: "--verbose",
            description: "Enable verbose logging.");
        verboseOption.AddAlias("-v");
        var adobeDirectoryOption = new Option<DirectoryInfo>
            (name: "--adobe-dir",
            description: "Path to the Adobe \"CoreSync\" directory.",
            getDefaultValue: () =>
            {
                return new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Adobe\CoreSync");
            })
        {
            ArgumentHelpName = "PATH",
        };
        var outputDirectoryOption = new Option<DirectoryInfo>
            (name: "--output-dir",
            description: "Path to the output directory. Will create a new directory if it doesn't already exist.",
            getDefaultValue: () =>
            {
                return new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Adobe\Fonts");
            })
        {
            ArgumentHelpName = "PATH"
        };

        rootCommand.Add(copyCommand);

        copyCommand.AddOption(dryRunOption);
        copyCommand.AddOption(verboseOption);
        copyCommand.AddOption(adobeDirectoryOption);
        copyCommand.AddOption(outputDirectoryOption);

        copyCommand.SetHandler(CopyFiles, dryRunOption, verboseOption, adobeDirectoryOption, outputDirectoryOption);

        return await rootCommand.InvokeAsync(args);
    }
    
    static void CopyFiles(bool isDry, bool isVerbose, DirectoryInfo adobeDirectory, DirectoryInfo outputDirectory)
    {
        DirectoryInfo coreSyncDirectory = adobeDirectory.GetDirectories("r", SearchOption.AllDirectories)[0];

        if (isVerbose)
        {
            Console.WriteLine($"Found CoreSync folder at {coreSyncDirectory}.");
        }

        if (!outputDirectory.Exists)
        {
            if (!isDry)
            {
                outputDirectory.Create();
            }

            if (isVerbose)
            {
                Console.WriteLine($"Created output directory at {outputDirectory}");
            }
        } else {
            if (isVerbose)
            {
                Console.WriteLine($"Found output directory at {outputDirectory}");
            }
        }

        FileInfo[] fontList = coreSyncDirectory.GetFiles();

        if (isVerbose)
        {
            Console.WriteLine($"Found {fontList.Length} fonts.");
        }

        foreach (FileInfo fontFile in fontList)
        {
            PrivateFontCollection fontCollection = new();
            fontCollection.AddFontFile(fontFile.FullName);

            if (isVerbose)
            {
                Console.WriteLine($"Processing file {fontFile}.");
            }

            string fontName = fontCollection.Families[0].Name;
            string outputPath = Path.Join(outputDirectory.FullName, fontName);
            string outputFileName = Path.ChangeExtension(outputPath, ".otf");

            if (isVerbose)
            {
                Console.WriteLine
                ($"Found font {fontName}.");
            }

            if (!isDry)
            {
                fontFile.CopyTo(outputFileName, true);

                if (isVerbose)
                {
                    Console.WriteLine($"Copied font {fontName} to {outputFileName}.");
                }
            }

            string[] splitFontPath = fontFile.FullName.Split("\\");
            string fontObfuscatedName = splitFontPath[^1];
            Console.WriteLine($"{fontObfuscatedName}\t->\t{fontName}.otf");

            fontCollection.Dispose();
        }
    }
}