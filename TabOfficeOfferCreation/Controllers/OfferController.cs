using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Identity.Web;
using PnP.Core.Auth;
using PnP.Core.Services;
using PnP.Core.Services.Builder.Configuration;
using TabOfficeOfferCreation.Model;

namespace TabOfficeOfferCreation.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class OfferController : ControllerBase
    {
        private readonly GraphServiceClient _graphClient;
        private readonly ITokenAcquisition _tokenAcquisition;
        private readonly IPnPContextFactory _pnpContextFactory;
        private readonly ILogger<GraphController> _logger;
        private readonly PnPCoreOptions _pnpCoreOptions;
        private string SiteUrl = "https://your-tenant.sharepoint.com/sites/Offerings"; // ToDo
        public OfferController(IPnPContextFactory pnpContextFactory, ITokenAcquisition tokenAcquisition, GraphServiceClient graphClient, ILogger<GraphController> logger,
            IOptions<PnPCoreOptions> pnpCoreOptions)
        {
            _tokenAcquisition = tokenAcquisition;
            _graphClient = graphClient;
            _pnpContextFactory = pnpContextFactory;
            _logger = logger;
            _pnpCoreOptions = pnpCoreOptions?.Value;
        }
        [HttpPost]
        public async Task<ActionResult<string>> Post(Offer offer)
        {
            string userID = User.GetObjectId(); //   Claims["preferred_username"];
            _logger.LogInformation($"Received from user {userID} with name {User.GetDisplayName()}");
            _logger.LogInformation($"Received Offer {offer.Title} with descr {offer.Description}");

            SPOController spoCtrl = new SPOController(_tokenAcquisition, _pnpContextFactory, _logger, _pnpCoreOptions);
            string result = await spoCtrl.CreateOfferFromTemplate(offer);
            return result;
        }

        
    }
}
