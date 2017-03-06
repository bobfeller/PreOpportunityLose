using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System.Runtime.Serialization;


namespace ASG.CRM.PreOpportunityLosePlugin
{
    public class PreOpportunityLosePlugin : IPlugin
    {
        public IOrganizationService wService;

        public void Execute(IServiceProvider serviceProvider)
        {
            #region variables

            Guid guidOpportunityId = new Guid();
            List<Quote> quoteToBeCanceled = new List<Quote>();

            #endregion

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));


            if (context.InputParameters.Contains("OpportunityClose") && context.InputParameters["OpportunityClose"] is Entity)
            {
                // Get a refrence to CRM API Services
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                wService = serviceFactory.CreateOrganizationService(context.UserId);

                //create entity context
                Entity entity = (Entity)context.InputParameters["OpportunityClose"];

                if (entity.LogicalName != OpportunityClose.EntityLogicalName)
                {
                    return;
                }
                try
                {
                    //create target entity as early bound
                    OpportunityClose TargetEntity = entity.ToEntity<OpportunityClose>();

                    //get the Opportunity Id
                    guidOpportunityId = TargetEntity.OpportunityId.Id;

                    //get all quote related to this Opportunity
                    quoteToBeCanceled = GetQuoteToBeCanceled(guidOpportunityId);

                    // Make sure there is not a won quote
                    if (quoteToBeCanceled != null)
                    { 
                        foreach (Quote quote in quoteToBeCanceled)
                        {
                            if (quote.StateCode == QuoteState.Won)
                            {
                                throw new InvalidPluginExecutionException("Unable to close this opportunity. There are Won quotes.");
                            }
                        }
                    }

                    if (quoteToBeCanceled != null)
                    {
                        foreach (Quote quote in quoteToBeCanceled)
                        {
                            if (quote.StateCode.Value == QuoteState.Closed)
                            {
                                continue;
                            }

                            if (quote.StateCode.Value == QuoteState.Draft)
                            {
                                //have to be activated first
                                ActivateDraftQuote(wService, quote);
                            }

                            //Execute Close as Lost
                            ExecuteCloseQuoteAsLost(wService, quote, quote.QuoteNumber, quote.StatusCode.Value);
                        }
                    }
                   
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }

        // Return all of the Quote from related Opportunity to be Canceled/Lost
        public List<Quote> GetQuoteToBeCanceled(Guid guidOpportunityId)
        {
            Quote quote = new Quote();
            //get the Quote to be canceled in same Opportunity, exclude the wonQuote
            myServiceContext context = new myServiceContext(wService);  //This comes from GeneratedCodeWithContext.cs
            var quoteToBeCanceled = from x in context.QuoteSet
                                    orderby x.CreatedOn descending
                                    where x.OpportunityId.Id == guidOpportunityId
                                    select x;

            if (quoteToBeCanceled.ToList().Count > 0)
                return quoteToBeCanceled.ToList<Quote>();
            else
                return null;
        }

        // Activate the Draft Quote
        public void ActivateDraftQuote(IOrganizationService service, Quote erQuote)
        {
            // Activate the quote
            SetStateRequest activateQuote = new SetStateRequest()
            {
                EntityMoniker = erQuote.ToEntityReference(),
                State = new OptionSetValue((int)QuoteState.Active),
                Status = new OptionSetValue((int)2) //in progress
            };
            service.Execute(activateQuote);
        }

        // Close the Quotes as Canceled
        public void ExecuteCloseQuoteAsCanceled(IOrganizationService service, Quote erQuote, string strQuoteNumber)
        {
            CloseQuoteRequest closeQuoteRequest = new CloseQuoteRequest()
            {
                QuoteClose = new QuoteClose()
                {
                    Subject = String.Format("Quote Closed (Canceled) - {0} - {1}", strQuoteNumber, DateTime.Now.ToString()),
                    QuoteId = erQuote.ToEntityReference()
                },
                Status = new OptionSetValue(6)
            };
            service.Execute(closeQuoteRequest);
        }
        // Close the Quotes as Opportunity Lost
        public void ExecuteCloseQuoteAsLost(IOrganizationService service, Quote erQuote, string strQuoteNumber, int statusCode)
        {
            int status = (statusCode == 9) ? 7 : 100000002;     // 7 - Expired, 10000002 - Closed Opp Lost

            CloseQuoteRequest closeQuoteRequest = new CloseQuoteRequest()
            {
                QuoteClose = new QuoteClose()
                {
                    Subject = String.Format("Quote Closed (Opportunity Lost) - {0} - {1}", strQuoteNumber, DateTime.Now.ToString()),
                    QuoteId = erQuote.ToEntityReference()
                },
                Status = new OptionSetValue(status)  // Opportunity Lost
            };
            service.Execute(closeQuoteRequest);
        }
    }
}