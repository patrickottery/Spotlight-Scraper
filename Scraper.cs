#region Main

Console.WriteLine("(*** Spotlight scraper beginning");
var destinationFolder = args.Length == 0 ? @"C:\Temp\Spotlight Images\" : args[0];
var destinationFolderFullPath = new DirectoryInfo(destinationFolder).FullName;
Console.WriteLine("(*** Destination Folder: "+ destinationFolder);
CreateFolderIfNotExists(destinationFolderFullPath);
var jsonMetaDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),  @"Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\TargetedContentCache\v3\338387");
var jsonMetaDataFiles = System.IO.Directory.GetFiles(jsonMetaDataFolder, "*.");
var spotlightImageInformation = GetSpotlightImageInformation(jsonMetaDataFiles);

foreach (var item in spotlightImageInformation)
{
    CopyFile(item);
}
Console.WriteLine("(*** Spotlight scraper finished");
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
        var fileName = GetFilename(json["properties"]["landscapeImage"]["image"].ToString());
        var destinationFileName = string.Format("{0} - {1}.jpg", fileName, SanitiseStringForFilename(description));

        list.Add(new SpotlightImageInformation
        {
            Description = description,
            LandscapeImageFullPath = landscapeImageFullPath,
            DestinationFileName = destinationFileName
        });
    }
    return list;
}

static string GetFilename(string fullFilePath)
{
    return Path.GetFileNameWithoutExtension(new FileInfo(fullFilePath).FullName);
}

static string SanitiseStringForFilename(string input)
{
   string invalidChars = System.Text.RegularExpressions.Regex.Escape( new string( System.IO.Path.GetInvalidFileNameChars() ) );
   string invalidRegStr = string.Format( @"([{0}]*\.+$)|([{0}]+)", invalidChars );

   return System.Text.RegularExpressions.Regex.Replace(input, invalidRegStr, "_" );
}

static Newtonsoft.Json.Linq.JObject GetJSONInfoFromFile(string filename)
{
    string rawJson = File.ReadAllText(filename);
    var json =  Newtonsoft.Json.Linq.JObject.Parse(rawJson);
    
    return json;
}

static void CreateFolderIfNotExists(string destinationFolderFullPath)
{
    if (!System.IO.Directory.Exists(destinationFolderFullPath))
    {
        try
        {
            Console.WriteLine("Creating destination folder: " + destinationFolderFullPath);
            System.IO.Directory.CreateDirectory(destinationFolderFullPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine(string.Format("(*** Couldn't create new folder <{0}> because of exception <{1}>", destinationFolderFullPath, ex.Message));
        }
    }
}

void CopyFile(SpotlightImageInformation item)
{
    var destinationFile = new FileInfo(destinationFolderFullPath + item.DestinationFileName);
    if (destinationFile.Exists)
    {
        Console.WriteLine("File already exists: " + destinationFile.FullName);
    }
    else
    {
        Console.WriteLine("Copying File: " + destinationFile.FullName);
        File.Copy(item.LandscapeImageFullPath, destinationFile.FullName);
    }
}

#endregion
public class SpotlightImageInformation
{
    public string Description { get; set; }
    public string LandscapeImageFullPath { get; set; }
    public string DestinationFileName { get; set; }
}