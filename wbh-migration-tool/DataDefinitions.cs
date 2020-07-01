using System;
using System.Collections.Generic;

namespace PRGX.Panoptic_Migration_Core
{
    public class VendorMaster
    {
        public string appVendorId { get; set; }
        public string vendorName { get; set; }
        public string vendorNumber { get; set; }

        public string vendorNumber2 { get; set; }

        public Dictionary<string, string> vendorAccounts;
        public VendorMaster()
        {
            vendorAccounts = new Dictionary<string, string>();
        }
    }

    public class CreditStatus
    {
        public string id { get; set; }
    }

    public class ProblemSubType
    {
        public string id { get; set; }
    }

    public class ProblemType
    {
        public string id { get; set; }
    }

    public class VendorAccountEntity
    {
        public object appVendorAccountId { get; set; }
    }

    public class Vendor
    {
        public string appVendorId { get; set; }
        public int vendorId { get; set; }
        public string vendorName { get; set; }
        public object vendorName2 { get; set; }
        public object vendorName3 { get; set; }
        public object vendorName4 { get; set; }
        public string vendorNumber { get; set; }
        public object vendorNumber2 { get; set; }
        public object vendorDescription { get; set; }
        public string vendorGroup { get; set; }
        public object businessUnit { get; set; }
        public object vendorLanguage { get; set; }
        public string vendorType { get; set; }
        public object companyCode { get; set; }
        public object taxId1 { get; set; }
        public string taxId2 { get; set; }
        public object inactiveDate { get; set; }
        public string auditVendorType { get; set; }
        public object custom { get; set; }
        public bool isGroupVendor { get; set; }
        public object vatReg { get; set; }
        public string sourceVenID { get; set; }
        public string uniqueVenID { get; set; }
        public string createdBy { get; set; }
        public DateTime createdDate { get; set; }
        public string lastModifiedBy { get; set; }
        public DateTime lastModifiedDate { get; set; }
    }

    public class VendorCreditRequest
    {
        public VendorCreditRequest()
        {
            vendor = new Vendor();
            creditStatus = new CreditStatus();
            problemType = new ProblemType();
        }
        public int id { get; set; }
        public int projectId { get; set; }
        public string parentGroupName { get; set; }
        public object creditReferenceNumber { get; set; }
        //public object Creditreferencenumber2 { get; set; }  //Custom
        public DateTime creditDate { get; set; }
        public Decimal creditAmount { get; set; }
        public string creditCurrency { get; set; }
        public string poNumber { get; set; }
        public CreditStatus creditStatus { get; set; }
        public ProblemSubType problemSubType { get; set; }
        public ProblemType problemType { get; set; }
        public object assignTo { get; set; }
        public object remindDate { get; set; }
        public string createdBy { get; set; }
        public DateTime createdDate { get; set; }
        public string lastModifiedBy { get; set; }
        public DateTime lastModifiedDate { get; set; }
        public Vendor vendor { get; set; }
        public VendorAccountEntity vendorAccountEntity { get; set; }
        //public DateTime Vendorapprovaldate { get; set; }        //Custom
    }

    public class CommentRequest
    {
        public bool archived { get; set; }
        public object parentComment { get; set; }
        public object commentThreadId { get; set; }
        public object readStatus { get; set; }
        public string content { get; set; }
        public int visLevel { get; set; }
        public DateTime createdDate { get; set; }
        public DateTime lastModifiedDate { get; set; }
        public string createdBy { get; set; }
        public string lastModifiedBy { get; set; }
        public CommentRequest()
        {
            archived = false;
        }
    }

    public class ProjectVendor
    {
        public long projectVendorId { get; set; }
        public long projectVendorStatusId { get; set; }
    }

    public class VendorSolicitationRequest
    {
        public long actionTakenValueId { get; set; }
        public object assignToId { get; set; }
        public object followUpDate { get; set; }
        public ProjectVendor projectVendor { get; set; }
        public DateTime createdDate { get; set; }
        public DateTime lastModifiedDate { get; set; }
        public string createdBy { get; set; }
        public string lastModifiedBy { get; set; }
        public VendorSolicitationRequest()
        {
            projectVendor = new ProjectVendor();
        }
    }

    public class FileMetadata
    {
        public string fileName { get; set; }
        public string aliasName { get; set; }
        public int docType { get; set; }
        public int fileOrder { get; set; }
        public string visibilityLevel { get; set; }
        public int moduleId { get; set; }
        public string objectId { get; set; }
        public object objectValueId { get; set; }
        public string appVendorId { get; set; }
        public int projectId { get; set; }
        public DateTime createdDate { get; set; }
        public DateTime lastModifiedDate { get; set; }
        public string createdBy { get; set; }
        public string lastModifiedBy { get; set; }
        public object projectVendorId { get; set; }
    }

    public class ProblemTypeId
    {
        public string id { get; set; }
        public string value { get; set; }
    }
    public class RootCauseId
    {
        public string id { get; set; }
        public string value { get; set; }
    }

    public class AdhocClaimRequest
    {
        public string vendorNumber { get; set; }
        public string vendorName { get; set; }
        public decimal claimAmount { get; set; }
        public decimal netClaimAmount { get; set; }
        public string claimCurrency { get; set; }
        public DateTime claimDate { get; set; }
        public string projectId { get; set; }
        public string vendorId { get; set; }
        public string clientClaimNumber { get; set; }
        public string businessUnit { get; set; }
        public string auditProject { get; set; }
        public string statusId { get; set; }
        public string status { get; set; }
        public string appVendorId { get; set; }
        public string auditName { get; set; }
        //public DateTime Deductdate { get; set; }        //Custom
        //public object CCAauditnumber { get; set; }        //Custom
        //public object Invoicenumber { get; set; }        //Custom
        //public DateTime Invoicedate { get; set; }        //Custom
        //public object PS_project_id { get; set; }        //Custom
        public DateTime createdDate { get; set; }
        public string createdBy { get; set; }
        public ProblemTypeId problemTypeId { get; set; }
        public RootCauseId rootCauseId { get; set; }
        public int templateId { get; set; }
    }

    public class ClaimInput
    {
        public string vendorName { get; set; }
        public object vendorNumber { get; set; }
        public object Vendornumber2 { get; set; }
        public string claimDate { get; set; }
        public string claimCurrency { get; set; }
        public Decimal claimAmount { get; set; }
        public object exchangeRate { get; set; }
        public object Processsolution { get; set; }
        public object Clientcompanynumber { get; set; }
        public object Clientcostcenter { get; set; }
        public object Clientplantcode { get; set; }
        public object Divisionmailstop { get; set; }
        public object Workarea { get; set; }
        public string auditName { get; set; }
        public string appVendorId { get; set; }
        //public object Invoicenumber { get; set; }       //Cusotm
        //public object PS_project_id { get; set; }       //Cusotm
        public ProblemTypeId problemTypeId { get; set; }
        public RootCauseId rootCauseId { get; set; }
    }

    public class CreditClaimRequest
    {
        public int templateId { get; set; }
        public string projectName { get; set; }
        public int projectId { get; set; }
        public IList<int> uniqueIds { get; set; }
        public ClaimInput claimInput { get; set; }
        public string source { get; set; }
    }

    public class WorkflowInput
    {
        public string userAction { get; set; }
        public int userActionId { get; set; }
    }

    public class WFActionRequest
    {
        public WorkflowInput workflowInput { get; set; }
        public string approvalNotes { get; set; }
        public WFActionRequest()
        {
            workflowInput = new WorkflowInput();
        }
    }

    public class DispositionId
    {
        public string id { get; set; }
        public string value { get; set; }
    }
    public class StatusId
    {
        public string id { get; set; }
        public string value { get; set; }
    }

    public class ClaimHeader
    {
        public DispositionId dispositionId { get; set; }
        public StatusId statusId { get; set; }
        public ClaimHeader()
        {
            dispositionId = new DispositionId();
            statusId = new StatusId();
        }
    }

    public class WFChangeStatusRequest
    {
        public ClaimHeader claimHeader { get; set; }
        public WorkflowInput workflowInput { get; set; }
        public string approvalNotes { get; set; }
        public WFChangeStatusRequest()
        {
            claimHeader = new ClaimHeader();
            workflowInput = new WorkflowInput();
        }
    }

    public enum StatusCode { Inprocess, Complete, PartiallyComplete, Error }

    public class APIReturnValue
    {
        public string PanopticId { get; set; }
        public StatusCode Status { get; set; }
        public string StatusDesc { get; set; }
    }
}
