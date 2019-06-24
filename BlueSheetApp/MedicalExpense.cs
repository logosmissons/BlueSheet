using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlueSheetApp
{

    public class SortedField
    {
        public enum EnumSorted { NotSorted, SortedAsc, SortedDesc };

        public String Field = String.Empty;
        public EnumSorted Sorted = EnumSorted.NotSorted;

        public SortedField()
        {
            Field = String.Empty;
            Sorted = EnumSorted.NotSorted;
        }
    }

    public class MedicalExpense
    {
        public double? BillAmount;
        public double? MemberDiscount;
        public double? CMMDiscount;
        public double? PersonalResponsibility;
        public double? CMMProviderPayment;
        public double? PastCMMProviderPayment;
        public double? PastReimbursement;
        public double? Reimbursement;
        public double? Balance;

        public MedicalExpense()
        {
            BillAmount = 0;
            MemberDiscount = 0;
            CMMDiscount = 0;
            PersonalResponsibility = 0;
            CMMProviderPayment = 0;
            PastCMMProviderPayment = 0;
            PastReimbursement = 0;
            Reimbursement = 0;
            Balance = 0;
        }

        public MedicalExpense(double? billAmount,
                              double? memberDiscount,
                              double? cmmDiscount,
                              double? personalResponsibility,
                              double? cmmProviderPayment,
                              double? pastCMMProviderPayment,
                              double? pastReimbursement,
                              double? reimbursement,
                              double? balance)
        {
            BillAmount = billAmount;
            MemberDiscount = memberDiscount;
            CMMDiscount = cmmDiscount;
            PersonalResponsibility = personalResponsibility;
            CMMProviderPayment = cmmProviderPayment;
            PastCMMProviderPayment = pastCMMProviderPayment;
            PastReimbursement = pastReimbursement;
            Reimbursement = reimbursement;
            Balance = balance;
        }
    }

 

    public class MedicalExpenseIneligible
    {
        public double? BillAmount;
        public double? AmountIneligible;

        public MedicalExpenseIneligible()
        {
            BillAmount = 0;
            AmountIneligible = 0;
        }

        public MedicalExpenseIneligible(double? billAmount,
                                        double? amountIneligible)
        {
            BillAmount = billAmount;
            AmountIneligible = amountIneligible;
        }
    }
    
    public class MedicalExpensePartiallyIneligible
    {
        public String INCD;
        public String PatientName;
        public String MedBill;
        public DateTime? ServiceDate;
        public DateTime? ReceiveDate;
        public String MedicalProvider;
        public Double? BillAmount;
        public Double? IneligibleAmount;
        public String IneligibleReason;

        public MedicalExpensePartiallyIneligible()
        {
            INCD = String.Empty;
            PatientName = String.Empty;
            MedBill = String.Empty;
            ServiceDate = null;
            ReceiveDate = null;
            MedicalProvider = String.Empty;
            BillAmount = 0;
            IneligibleAmount = 0;
            IneligibleReason = String.Empty;
        }
    }

    public class CMMPendingPayment
    {
        public double? BillAmount;
        public double? MemberDiscount;
        public double? CMMDiscount;
        public double? PersonalResponsibility;
        public double? SharedAmount;
        public double? AmountWillBeShared;

        public CMMPendingPayment()
        {
            BillAmount = 0;
            MemberDiscount = 0;
            CMMDiscount = 0;
            PersonalResponsibility = 0;
            SharedAmount = 0;
            AmountWillBeShared = 0;
        }

        public CMMPendingPayment(double? billAmount,
                                 double? memberDiscount,
                                 double? cmmDiscount,
                                 double? personalResponsibility,
                                 double? sharedAmount,
                                 double? amountWillBeShared)
        {
            BillAmount = billAmount;
            MemberDiscount = memberDiscount;
            CMMDiscount = cmmDiscount;
            PersonalResponsibility = personalResponsibility;
            SharedAmount = sharedAmount;
            AmountWillBeShared = amountWillBeShared;
        }
    }

    public class Pending
    {
        public double? BillAmount;
        public double? Balance;
        public double? MemberDiscount;
        public double? CMMDiscount;
        public double? SharedAmount;
        public double? PendingAmount;
        
        public Pending()
        {
            BillAmount = 0;
            Balance = 0;
            MemberDiscount = 0;
            CMMDiscount = 0;
            SharedAmount = 0;
            PendingAmount = 0;
        }

        public Pending(double? billAmount,
                       double? balance,
                       double? memberDiscount,
                       double? cmmDiscount,
                       double? sharedAmount,
                       double? pendingAmount)
        {
            BillAmount = billAmount;
            Balance = balance;
            MemberDiscount = memberDiscount;
            CMMDiscount = cmmDiscount;
            SharedAmount = sharedAmount;
            PendingAmount = pendingAmount;
        }
    }

    public class PaidMedicalExpenseTableRow
    {
        //public String CheckNo;
        //public String Issue_Date;
        //public String INCD;
        public String PatientName;
        public String MED_BILL;
        public DateTime? Bill_Date;
        public String Medical_Provider;
        public String Bill_Amount;
        public String Member_Discount;
        public String CMM_Discount;
        public String Personal_Responsibility;
        public String CMM_Provider_Payment;
        public String PastCMM_Provider_Payment;
        public String PastReimbursement;
        public String Reimbursement;
        public String Balance;

        public PaidMedicalExpenseTableRow()
        {
            //CheckNo = String.Empty;
            //Issue_Date = String.Empty;
            PatientName = String.Empty;
            MED_BILL = String.Empty;
            Bill_Date = null;
            Medical_Provider = String.Empty;
            Bill_Amount = String.Empty;
            Member_Discount = String.Empty;
            CMM_Discount = String.Empty;
            Personal_Responsibility = String.Empty;
            CMM_Provider_Payment = String.Empty;
            PastCMM_Provider_Payment = String.Empty;
            PastReimbursement = String.Empty;
            Reimbursement = String.Empty;
            Balance = String.Empty;
        }
    }

    public class CMMPendingPaymentTableRow
    {
        //public String INCD;
        public String PatientName;
        public String MED_BILL;
        public String Bill_Date;
        public String Due_Date;
        public String Medical_Provider;
        public String Bill_Amount;
        public String Member_Discount;
        public String CMM_Discount;
        public String PersonalResponsibility;
        public String Shared_Amount;
        public String Balance;

        public CMMPendingPaymentTableRow()
        {
            PatientName = String.Empty;
            MED_BILL = String.Empty;
            Bill_Date = String.Empty;
            Due_Date = String.Empty;
            Medical_Provider = String.Empty;
            Bill_Amount = String.Empty;
            Member_Discount = String.Empty;
            CMM_Discount = String.Empty;
            PersonalResponsibility = String.Empty;
            Shared_Amount = String.Empty;
            Balance = String.Empty;
        }
    }

    public class PendingTableRow
    {
        //public String INCD;
        public String PatientName;
        public String MED_BILL;
        public String Bill_Date;
        public String Due_Date;
        public String Medical_Provider;
        public String Bill_Amount;
        public String Balance;
        public String Member_Discount;
        public String CMM_Discount;
        public String Shared_Amount;
        public String Pending_Reason;

        public PendingTableRow()
        {
            PatientName = String.Empty;
            MED_BILL = String.Empty;
            Bill_Date = String.Empty;
            Due_Date = String.Empty;
            Medical_Provider = String.Empty;
            Bill_Amount = String.Empty;
            Balance = String.Empty;
            Member_Discount = String.Empty;
            CMM_Discount = String.Empty;
            Shared_Amount = String.Empty;
            Pending_Reason = String.Empty;
        }
    }

    public class BillIneligibleTableRow
    {
        //public String INCD;
        public String PatientName;
        public String MED_BILL;
        public String Bill_Date;
        public String Received_Date;
        public String Medical_Provider;
        public String Bill_Amount;
        public String Amount_Ineligible;
        public String Ineligible_Reason;

        public BillIneligibleTableRow()
        {
            PatientName = String.Empty;
            MED_BILL = String.Empty;
            Bill_Date = String.Empty;
            Received_Date = String.Empty;
            Medical_Provider = String.Empty;
            Bill_Amount = String.Empty;
            Amount_Ineligible = String.Empty;
            Ineligible_Reason = String.Empty;
        }
    }

    public class BillIneligibleRow
    {
        public Double? Bill_Amount;
        public Double? Amount_Ineligible;

        public BillIneligibleRow()
        {
            Bill_Amount = 0;
            Amount_Ineligible = 0;
        }
    }

    public class Incident : IEquatable<Incident>, IComparable<Incident>
    {
        public String Name { get; set; }
        public String PatientName { get; set; }
        public String ICD10_Code { get; set; }

        public Incident()
        {
            Name = String.Empty;
            PatientName = String.Empty;
            ICD10_Code = String.Empty;
        }
        public Incident(String name, String patientName, String icd10_code)
        {
            Name = name;
            PatientName = patientName;
            ICD10_Code = icd10_code;
        }

        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            Incident objIncd = obj as Incident;
            if (objIncd == null) return false;
            else return Equals(objIncd);
        }

        public bool Equals(Incident incd)
        {
            if (incd == null) return false;
            return (this.Name.Equals(incd.Name));
        }

        public int SortByNameAscending(String Name1, String Name2)
        {
            return Name1.CompareTo(Name2);
        }

        public int CompareTo(Incident compareIncd)
        {
            if (compareIncd == null)
                return 1;
            else
                return this.Name.CompareTo(compareIncd.Name);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }

    public class CheckInfo
    {
        public String CheckNumber { get; set; }
        public DateTime dtCheckIssueDate { get; set; }
        public Double? CheckAmount { get; set; }
        public String PaidTo { get; set; }

        public CheckInfo()
        {
            CheckNumber = String.Empty;
            dtCheckIssueDate = DateTime.Today;
            CheckAmount = 0;
            PaidTo = String.Empty;
                
        }

        public CheckInfo(String checkno, DateTime issuedate, Double amount, String paidTo)
        {
            CheckNumber = checkno;
            dtCheckIssueDate = issuedate;
            CheckAmount = amount;
            PaidTo = paidTo;
        }
    }

    public class ACHInfo
    {
        public String ACHNumber { get; set; }
        public DateTime dtACHDate { get; set; }
        public Double? ACHAmount { get; set; }
        public String PaidTo { get; set; }

        public ACHInfo()
        {
            ACHNumber = String.Empty;
            dtACHDate = DateTime.Today;
            ACHAmount = 0;
            PaidTo = String.Empty;
        }

        public ACHInfo(String ach_no, DateTime ach_date, Double amount, String paidTo)
        {
            ACHNumber = ach_no;
            dtACHDate = ach_date;
            ACHAmount = amount;
            PaidTo = paidTo;
        }
    }

    public class CreditCardPaymentInfo
    {
        public DateTime dtPaymentDate { get; set; }
        public Double? CCPaymentAmount { get; set; }
        public String PaidTo { get; set; }

        public CreditCardPaymentInfo()
        {
            dtPaymentDate = DateTime.Today;
            CCPaymentAmount = 0;
            PaidTo = String.Empty;
        }

        public CreditCardPaymentInfo(DateTime cc_payment_date, Double amount, String paidTo)
        {
            dtPaymentDate = cc_payment_date;
            CCPaymentAmount = amount;
            PaidTo = paidTo;
        }
    }

    public class PersonalResponsibilityTotalInfo
    {
        public String IncidentNo;
        public String ICD10CodeDescription;
        public DateTime? IncidentOccurrenceDate;
        public Decimal PersonalResponsibilityTotal;

        public PersonalResponsibilityTotalInfo()
        {
            IncidentNo = String.Empty;
            ICD10CodeDescription = String.Empty;
            IncidentOccurrenceDate = null;
            PersonalResponsibilityTotal = 0;
        }

        public PersonalResponsibilityTotalInfo(String incident_no, DateTime incident_occurrence_date, Decimal personal_responsibility_total)
        {
            IncidentNo = incident_no;
            IncidentOccurrenceDate = incident_occurrence_date;
            PersonalResponsibilityTotal = personal_responsibility_total;
        }

        public PersonalResponsibilityTotalInfo(String incident_no, String icd10code_description, DateTime incident_occurrence_date, Decimal personal_responsibility_total)
        {
            IncidentNo = incident_no;
            ICD10CodeDescription = icd10code_description;
            IncidentOccurrenceDate = incident_occurrence_date;
            PersonalResponsibilityTotal = personal_responsibility_total;
        }
    }

    public class PersonalResponsibilityInfo
    {
        public String MedBillName;
        public DateTime? BillDate;
        public String MedicalProvider;
        public Double BillAmount;
        public String Type;
        public String PersonalResponsibilityType;
        public Double? MemberPayment;
        public Double? MemberDiscount;
        public Double? ThirdPartyDiscount;
        public Double PersonalResponsibilityTotal;

        public PersonalResponsibilityInfo()
        {
            MedBillName = String.Empty;
            BillDate = null;
            MedicalProvider = String.Empty;
            BillAmount = 0;
            Type = String.Empty;
            PersonalResponsibilityType = String.Empty;
            MemberPayment = null;
            MemberDiscount = null;
            ThirdPartyDiscount = null;
            PersonalResponsibilityTotal = 0;
        }

        public PersonalResponsibilityInfo(String medbill_name, 
                                          DateTime bill_date, 
                                          String medical_provider, 
                                          Double bill_amount, 
                                          String type, 
                                          String personal_responsibility_type, 
                                          Double member_payment,
                                          Double member_discount,
                                          Double third_party_discount,
                                          Double personal_responsibility_total)
        {
            MedBillName = medbill_name;
            BillDate = bill_date;
            MedicalProvider = medical_provider;
            BillAmount = bill_amount;
            Type = type;
            PersonalResponsibilityType = personal_responsibility_type;
            MemberPayment = member_payment;
            MemberDiscount = member_discount;
            ThirdPartyDiscount = third_party_discount;
            PersonalResponsibilityTotal = personal_responsibility_total;
        }
    }

    public class SettlementIneligibleInfo
    {
        public String MedBillName;
        public DateTime? BillDate;
        public String MedicalProvider;
        public Double BillAmount;
        public String Type;
        public Double? IneligibleAmount;
        public String IneligibleReason;

        public SettlementIneligibleInfo()
        {
            MedBillName = String.Empty;
            BillDate = null;
            MedicalProvider = String.Empty;
            BillAmount = 0;
            Type = String.Empty;
            IneligibleAmount = 0;
            IneligibleReason = String.Empty;
        }

        public SettlementIneligibleInfo(String medbill_name,
                                        DateTime bill_date,
                                        String medical_provider,
                                        Double bill_amount,
                                        String type,
                                        Double ineligible_amount,
                                        String ineligible_reason)
        {
            MedBillName = medbill_name;
            BillDate = bill_date;
            MedicalProvider = medical_provider;
            BillAmount = bill_amount;
            Type = type;
            IneligibleAmount = ineligible_amount;
            IneligibleReason = ineligible_reason;
        }
    }

    public class PersonalResponsibilityExpense
    {
        public Double BillAmount;
        //public Double SettlementAmount;
        public Double MemberPayment;
        public Double MemberDiscount;
        public Double ThirdPartyDiscount;
        public Double PersonalResponsiblityTotal;

        public PersonalResponsibilityExpense()
        {
            BillAmount = 0;
            //SettlementAmount = 0;
            MemberPayment = 0;
            MemberDiscount = 0;
            ThirdPartyDiscount = 0;
            PersonalResponsiblityTotal = 0;
        }

        public PersonalResponsibilityExpense(Double bill_amount, Double member_payment, Double member_discount, Double third_party_discount, Double personal_responsibility_total)
        {
            BillAmount = bill_amount;
            //SettlementAmount = settlement_amount;
            MemberPayment = member_payment;
            MemberDiscount = member_discount;
            ThirdPartyDiscount = third_party_discount;
            PersonalResponsiblityTotal = personal_responsibility_total;
        }
    }
}
