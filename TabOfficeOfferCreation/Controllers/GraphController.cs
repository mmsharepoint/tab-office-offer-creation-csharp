using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Web;
using TabOfficeOfferCreation.Model;
using static System.Net.WebRequestMethods;

namespace TabOfficeOfferCreation.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class GraphController : ControllerBase
    {
        private readonly GraphServiceClient _graphClient;
        private readonly ITokenAcquisition _tokenAcquisition;
        private readonly ILogger<GraphController> _logger;
        private string SiteUrl = "https://your-tenant.sharepoint.com/sites/Offerings"; // ToDo
        public GraphController(ITokenAcquisition tokenAcquisition, GraphServiceClient graphClient, ILogger<GraphController> logger)
        {
            _tokenAcquisition = tokenAcquisition;
            _graphClient = graphClient;
            _logger = logger;
        }

        [HttpPost]
        public async Task<ActionResult<string>> Post(Offer offer)
        {            
            string userID = User.GetObjectId(); //   Claims["preferred_username"];
            _logger.LogInformation($"Received from user {userID} with name {User.GetDisplayName()}");
            _logger.LogInformation($"Received Offer {offer.Title} with descr {offer.Description}");

            SiteDrive siteDrive = await getDriveIdByUrl(SiteUrl);
            DriveItemResult docTemplateInfo = await getFileInfoName(siteDrive.SiteId, "DocumentTemplate.docx");
            DriveItem newFile = await createNewDocument(siteDrive.DriveId, $"{offer.Title}.docx", siteDrive.RootId, docTemplateInfo.Id);
            DriveItemResult newFileInfo = await getFileInfoName(siteDrive.SiteId, $"{offer.Title}.docx");
            bool update = await updateLibraryItem(siteDrive.SiteId, siteDrive.DriveId, offer, newFileInfo.Id);
            if (update)
            {
                return Ok(newFileInfo.WebUrl);
            }
            else
            {
                return BadRequest("Something went wrong");
            }
        }

        private async Task<SiteDrive> getDriveIdByUrl (string url)
        {
            Site site = await _graphClient.Sites["mmoellermvp.sharepoint.com:/teams/Offerings"].GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Expand = new string[] { "drive" };
            });
            DriveItem root = await _graphClient.Drives[site.Drive.Id].Root.GetAsync();
            return new SiteDrive { SiteId = site.Id, DriveId = site.Drive.Id, RootId = root.Id};
        }

        private async Task<DriveItemResult> getFileInfoName(string siteId, string fileName)
        {
            ListItemCollectionResponse docLibItems = await _graphClient.Sites[siteId].Lists["Documents"].Items.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.Headers.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly"); // Dirty, as Filename not indexed
                requestConfiguration.QueryParameters.Filter = $"fields/FileLeafRef eq '{fileName}'";
                requestConfiguration.QueryParameters.Expand = new string[] { "fields", "driveItem" };
            });
            DriveItem file = docLibItems.Value.First<ListItem>().DriveItem;
            return new DriveItemResult { Id = file.Id, WebUrl = file.WebUrl };
        }
        private async Task<DriveItem> createNewDocument(string driveId, string fileName, string rootId, string fileTemplateId)
        {
            var requestBody = new Microsoft.Graph.Drives.Item.Items.Item.Copy.CopyPostRequestBody
            {
                ParentReference = new ItemReference
                {
                    DriveId = driveId,
                    Id = rootId
                },
                Name = fileName
            };
            var result = await _graphClient.Drives[driveId].Items[fileTemplateId].Copy.PostAsync(requestBody);
            return result;
        }
        private async Task<bool> updateLibraryItem(string siteId, string driveId, Offer offer, string driveItemId)
        {
            ListItem listItem = await _graphClient.Drives[driveId].Items[driveItemId].ListItem.GetAsync();
            int listItemId;
            Int32.TryParse(listItem.Id, out listItemId);

            FieldValueSet fields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>
                    {                        
                        {
                            "Title",
                            offer.Title
                        },
                        {
                            "OfferingDescription",
                            offer.Description
                        },
                        {
                            "OfferingNetPrice",
                            offer.Price.ToString()
                        },
                        {
                            "OfferingDate",
                            offer.OfferDate.ToString("s")
                        },
                        {
                            "OfferingVAT",
                            offer.SelectedVAT
                        }
                    }
                };

            var result = await _graphClient.Sites[siteId].Lists["Documents"].Items[listItem.Id].Fields.PatchAsync(fields);
            return result.Id == listItem.Id;
        }
    }
}
