using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace Dynamics365_Plugins
{
    public class TestPlugin : IPlugin
    {

        public void Execute(IServiceProvider serviceProvider)
        {
            string address = "";

            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));


            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                try
                {
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    Entity entity = context.PostEntityImages["postImage"];
                    if (entity == null)
                        return;

                    Guid contactId = entity.GetAttributeValue<EntityReference>("customerid").Id;

                    ColumnSet columnSet = new ColumnSet();
                    columnSet.AddColumn("address1_composite");
                    columnSet.AddColumn("address2_composite");
                    columnSet.AddColumn("address3_composite");

                    Entity contact = service.Retrieve("contact", contactId, columnSet);

                    if (!contact.Contains("address1_composite"))
                    {
                        address = "address1";
                    }
                    else if(!contact.Contains("address2_composite"))
                    {
                        address = "address2";
                    }
                    else if (!contact.Contains("address3_composite"))
                    {
                        address = "address3";
                    }
                    if (address == null)
                        return;

                    Entity contactUpdate = new Entity("contact");
                    contactUpdate["contactid"] = contactId;

                    TryAssignValue(entity, contactUpdate, "billto_city", address + "_city");
                    TryAssignValue(entity, contactUpdate, "billto_country", address + "_country");
                    TryAssignValue(entity, contactUpdate, "billto_line1", address + "_line1");
                    TryAssignValue(entity, contactUpdate, "billto_line2", address + "_line2");
                    TryAssignValue(entity, contactUpdate, "billto_line3", address + "_line3");
                    TryAssignValue(entity, contactUpdate, "billto_stateorprovince", address + "_stateorprovince");
                    TryAssignValue(entity, contactUpdate, "billto_postalcode", address + "_postalcode");


                    tracingService.Trace("Updating Contact {0} Address.", contactId);
                    service.Update(contactUpdate);
                }

                catch(FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occured in TestPlugin", ex);
                }

                catch(Exception ex)
                {
                    tracingService.Trace("TestPlugin: {0}", ex.ToString());
                    throw;
                }
            }
        } 

        private void TryAssignValue(Entity src, in Entity dest, string srcAttribute, string destAttribute)
        {
            if(src.Contains(srcAttribute))
            {
                dest[destAttribute] = src[srcAttribute];
            }
        }
    }
}
