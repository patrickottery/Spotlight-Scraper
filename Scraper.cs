#region Using
using System.Security.Cryptography;
using Newtonsoft.Json.Linq;
#endregion

#region Main
Console.WriteLine("(*** Spotlight scraper beginning");
var destinationFolder = args.Length == 0 ? @"C:\Temp\Spotlight Images\" : args[0];
var destinationFolderFullPath = new DirectoryInfo(destinationFolder).FullName;
Console.WriteLine("(*** Destination Folder: "+ destinationFolderFullPath);
CreateFolderIfNotExists(destinationFolderFullPath);
var jsonMetaDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),  @"Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\TargetedContentCache\v3\338387");
var jsonMetaDataFiles = Directory.GetFiles(jsonMetaDataFolder, "*.");
var spotlightImageInformation = GetSpotlightImageInformation(jsonMetaDataFiles);

foreach (var item in spotlightImageInformation)
{
    CopyFile(item);
}
Console.WriteLine("*** Spotlight scraper finished");
#endregion


#region Methods
List<SpotlightImageInformation> GetSpotlightImageInformation(string[] jsonMetaDataFiles)
{
    List<SpotlightImageInformation> list = new List<SpotlightImageInformation>();
    foreach (var item in jsonMetaDataFiles)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        var json = GetJSONInfoFromFile(item);
        var description = json["items"][0]["properties"]["description"]["text"].ToString();
        var landscapeImageFullPath = json["properties"]["landscapeImage"]["image"].ToString();
        var fileName = GetFileNameWithoutExtension(json["properties"]["landscapeImage"]["image"].ToString());
        var destinationFileName = GetDestinationFilename(description) ;

        list.Add(new SpotlightImageInformation
        {
            Description = description,
            LandscapeImageFullPath = landscapeImageFullPath,
            DestinationFileName = destinationFileName
        });
    }
    return list;
}

string GetDestinationFilename(string description)
{
    string invalidChars = System.Text.RegularExpressions.Regex.Escape(new string(System.IO.Path.GetInvalidFileNameChars()));
    string invalidRegStr = string.Format(@"([{0}]*\.+$)|([{0}]+)", invalidChars);
    description = System.Text.RegularExpressions.Regex.Replace(description, invalidRegStr, "_");
    var fileName = string.Format("{0}.jpg", description);
    return fileName;
}

static bool FilesAreEqual(FileInfo first, FileInfo second)
{
    byte[] firstHash = MD5.Create().ComputeHash(first.OpenRead());
    byte[] secondHash = MD5.Create().ComputeHash(second.OpenRead());

    for (int i=0; i<firstHash.Length; i++)
    {
        if (firstHash[i] != secondHash[i])
            return false;
    }
    return true;
}


static string GetFileNameWithoutExtension(string fullFilePath)
{
    return Path.GetFileNameWithoutExtension(new FileInfo(fullFilePath).FullName);
}

static JObject GetJSONInfoFromFile(string filename)
{
    string rawJson = File.ReadAllText(filename);
    var json =  JObject.Parse(rawJson);
    
    return json;
}

static void CreateFolderIfNotExists(string destinationFolderFullPath)
{
    if (!Directory.Exists(destinationFolderFullPath))
    {
        try
        {
            Console.WriteLine("Creating destination folder: " + destinationFolderFullPath);
            Directory.CreateDirectory(destinationFolderFullPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine(string.Format("(*** Couldn't create new folder <{0}> because of exception <{1}>", destinationFolderFullPath, ex.Message));
        }
    }
}

void CopyFile(SpotlightImageInformation item)
{
    var destinationFileInfo = new FileInfo(destinationFolderFullPath + item.DestinationFileName);
    var sourceFileInfo = new FileInfo(item.LandscapeImageFullPath);
    var finished = false;
    while (!finished)
    {
        try
        {
            destinationFileInfo = new FileInfo(destinationFolderFullPath + item.DestinationFileName); //to ensure updated name used
            Console.WriteLine("Copying File: " + destinationFileInfo.FullName);
            File.Copy(item.LandscapeImageFullPath, destinationFileInfo.FullName);
            finished = true;
        }
        catch (IOException copyError)
        {
            Console.WriteLine("Copy Error: " + copyError.Message);
            if (FilesAreEqual(sourceFileInfo, destinationFileInfo))
            {
                Console.WriteLine("File already exists and is identical.");
                finished = true;
            }
            else
            {
                item.DestinationFileName = UpdateFileNameIfNeeded(item.DestinationFileName);
            }
        }
    }
}

string UpdateFileNameIfNeeded(string fileName)
{
    Console.WriteLine("Updating Filename: " + fileName);
    var originalFilename = fileName;
    var finished = false;
    var counter = 0;
    while (!finished)
    {
        var destinationFileInfo = new FileInfo(destinationFolderFullPath + fileName);
        if (!destinationFileInfo.Exists)
        {
            finished = true;
        }
        else
        {
            counter ++;
            fileName = originalFilename.Replace(".jpg", string.Format(" {0}.jpg", counter.ToString()));
        }
    }

    Console.WriteLine("Updated filename: " + fileName);
    return fileName;
}

#endregion
public class SpotlightImageInformation
{
    public string Description { get; set; }
    public string LandscapeImageFullPath { get; set; }
    public string DestinationFileName { get; set; }
}

