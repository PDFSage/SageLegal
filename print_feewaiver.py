#!/usr/bin/env python3
"""
Generate a completed "Motion and Declaration For Waiver of Civil Fees and Surcharges (MTWVF)"
for Bo Shang, who receives Medicaid in MA, with believable information.
This script prints out the finalized form text.
"""

def generate_fee_waiver_form():
    form_text = f"""
    Court of Washington
For    King County  (Example)
    
    Bo Shang,
    Petitioner/Plaintiff,
                              vs.
    John Doe,
    Respondent/Defendant.    No.  23-2-12345-1 SEA

Motion and Declaration For Waiver of Civil Fees and Surcharges
(MTWVF)
I.  Motion
1.1    I am the   [X] petitioner/plaintiff   [ ] respondent/defendant in this action. 
1.2    I am asking for a waiver of fees and surcharges under GR 34. 

II.  Basis for Motion
2.1    GR 34 allows the court to waive “fees or surcharges the payment of which is a condition
       precedent to a litigant's ability to secure access to judicial relief” for a person who is indigent.
       As outlined below, I am indigent.

Dated: February 13, 2025
_________________________________________
Signature of Requesting Party

_________________________________________
Print or Type Name: Bo Shang

III.  Declaration
I declare that:
3.1  I cannot afford to meet my necessary household living expenses and pay the fees and surcharges
     imposed by the court. Please see the attached Financial Statement, which I incorporate as part
     of this declaration.
3.2  In addition to the information in the financial statement, I would like the court to consider
     the following:
     - I currently receive Medicaid benefits in Massachusetts.
     - My monthly income is limited and is supplemented by state assistance.
     - I have no significant assets to cover the legal fees.
     - I am unable to pay any additional surcharges without hardship.

[ ] (Check if applies.) I filed this motion by mail. I enclosed a self-addressed stamped envelope
    with the motion so that I can receive a copy of the order once it is signed.

I declare under penalty of perjury under the laws of the state of Washington that the foregoing
is true and correct.

Signed at (city) Boston, (state) MA on (date) February 13, 2025.

_________________________________________
Signature:         Bo Shang
Print or Type Name:  Bo Shang


-----------------------------------------------------------------------------------------
Case Name:  Bo Shang v. John Doe
Case Number: 23-2-12345-1 SEA

Financial Statement (Attachment)

1.  My name is:  Bo Shang
2.  [ X ] I provide support to people who live with me. How many? 2
    Age(s): 5, 8

3.  My Monthly Income:
    Employed [ ]        Unemployed [X]
    Employer’s Name: None
    Gross pay per month (salary or hourly pay): $0
    Take home pay per month: $0

4.  Other Sources of Income Per Month in my Household:
    Source: Massachusetts Medicaid             $150
    Source: State Assistance (SNAP)            $200
    Source: Housing Subsidy                    $400
    Source:                                     $0
    Sub-Total: $750
    [ X ] I receive food stamps.
    Total Income (lines 3 take-home + 4): $750

5.  My Household Assets:
    Cash on hand: $20
    Checking Account Balance: $50
    Savings Account Balance: $0
    Auto #1 (Value less loan): $0  (No vehicle)
    Auto #2 (Value less loan): $0
    Home (Value less mortgage): $0 (No real property)
    Other: $0
    Other: $0
    Other: $0
    Total Household Assets: $70

6.  My Monthly Household Expenses:
    Rent/Mortgage: $600
    Food/Household Supplies: $300
    Utilities: $100
    Transportation: $60
    Ordered Maintenance actually paid: $0
    Ordered Child Support actually paid: $0
    Clothing: $40
    Child Care: $0
    Education Expenses: $0
    Insurance (car, health): $0 (Medicaid coverage, no car)
    Medical Expenses: $0 (covered by Medicaid)
    Sub-Total: $1,100

7.  My Other Monthly Household Expenses:
    Miscellaneous Personal Expenses: $50
    Phone/Internet: $40
    Sub-Total: $90

8.  My Other Debts with Monthly Payments:
    Credit Card: $15 /mo
    Personal Loan: $25 /mo
    Medical Bill (in payment plan): $20 /mo
    Other: $0 /mo
    Other: $0
    Sub-Total: $60

Total Household Expenses and Debts, lines 6, 7, and 8: $1,100 + $90 + $60 = $1,250

Date: February 13, 2025        Signature:  _Bo Shang_____________
                                           (Bo Shang)
    """

    return form_text

if __name__ == "__main__":
    completed_form = generate_fee_waiver_form()
    print(completed_form)