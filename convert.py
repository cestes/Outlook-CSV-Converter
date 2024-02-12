import sqlite3
import json

#=====================================================================================================================
#This follows the data structure that Outlook creates when you export to a CSV - I loaded it into a SQLite database
#for another purpose and just left it like this.
#=====================================================================================================================
class Contact(object): 
    def __init__(self,
        Title,
        FirstName,
        MiddleName,
        LastName,
        Suffix,
        Company,
        Department,
        JobTitle,
        BusinessStreet,
        BusinessStreet2,
        BusinessStreet3,
        BusinessCity,
        BusinessState,
        BusinessPostalCode,
        BusinessCountryRegion,
        HomeStreet,
        HomeStreet2,
        HomeStreet3,
        HomeCity,
        HomeState,
        HomePostalCode,
        HomeCountryRegion,
        OtherStreet,
        OtherStreet2,
        OtherStreet3,
        OtherCity,
        OtherState,
        OtherPostalCode,
        OtherCountryRegion,
        AssistantsPhone,
        BusinessFax,
        BusinessPhone,
        BusinessPhone2,
        Callback,
        CarPhone,
        CompanyMainPhone,
        HomeFax,
        HomePhone,
        HomePhone2,
        ISDN,
        MobilePhone,
        OtherFax,
        OtherPhone,
        Pager,
        PrimaryPhone,
        RadioPhone,
        TTYTDDPhone,
        Telex,
        Account,
        Anniversary,
        AssistantsName,
        BillingInformation,
        Birthday,
        BusinessAddressPOBox,
        Categories,
        Children,
        DirectoryServer,
        EmailAddress,
        EmailType,
        EmailDisplayName,
        Email2Address,
        Email2Type,
        Email2DisplayName,
        Email3Address,
        Email3Type,
        Email3DisplayName,
        Gender,
        GovernmentIDNumber,
        Hobby,
        HomeAddressPOBox,
        Initials,
        InternetFreeBusy,
        Keywords,
        Language,
        Location,
        ManagersName,
        Mileage,
        Notes,
        OfficeLocation,
        OrganizationalIDNumber,
        OtherAddressPOBox,
        Priority,
        Private,
        Profession,
        ReferredBy,
        Sensitivity,
        Spouse,
        User1,
        User2,
        User3,
        User4,
        WebPage):

        self.Title = Title
        self.FirstName = FirstName
        self.MiddleName = MiddleName
        self.LastName = LastName
        self.Suffix = Suffix
        self.Company = Company
        self.Department = Department
        self.JobTitle = JobTitle
        self.BusinessStreet = BusinessStreet
        self.BusinessStreet2 = BusinessStreet2
        self.BusinessStreet3 = BusinessStreet3
        self.BusinessCity = BusinessCity
        self.BusinessState = BusinessState
        self.BusinessPostalCode = BusinessPostalCode
        self.BusinessCountryRegion = BusinessCountryRegion
        self.HomeStreet = HomeStreet
        self.HomeStreet2 = HomeStreet2
        self.HomeStreet3 = HomeStreet3
        self.HomeCity = HomeCity
        self.HomeState = HomeState
        self.HomePostalCode = HomePostalCode
        self.HomeCountryRegion = HomeCountryRegion
        self.OtherStreet = OtherStreet
        self.OtherStreet2 = OtherStreet2
        self.OtherStreet3 = OtherStreet3
        self.OtherCity = OtherCity
        self.OtherState = OtherState
        self.OtherPostalCode = OtherPostalCode
        self.OtherCountryRegion = OtherCountryRegion
        self.AssistantsPhone = AssistantsPhone
        self.BusinessFax = BusinessFax
        self.BusinessPhone = BusinessPhone
        self.BusinessPhone2 = BusinessPhone2
        self.Callback = Callback
        self.CarPhone = CarPhone
        self.CompanyMainPhone = CompanyMainPhone
        self.HomeFax = HomeFax
        self.HomePhone = HomePhone
        self.HomePhone2 = HomePhone2
        self.ISDN = ISDN
        self.MobilePhone = MobilePhone
        self.OtherFax = OtherFax
        self.OtherPhone = OtherPhone
        self.Pager = Pager
        self.PrimaryPhone = PrimaryPhone
        self.RadioPhone = RadioPhone
        self.TTYTDDPhone = TTYTDDPhone
        self.Telex = Telex
        self.Account = Account
        self.Anniversary = Anniversary
        self.AssistantsName = AssistantsName
        self.BillingInformation = BillingInformation
        self.Birthday = Birthday
        self.BusinessAddressPOBox = BusinessAddressPOBox
        self.Categories = Categories
        self.Children = Children
        self.DirectoryServer = DirectoryServer
        self.EmailAddress = EmailAddress
        self.EmailType = EmailType
        self.EmailDisplayName = EmailDisplayName
        self.Email2Address = Email2Address
        self.Email2Type = Email2Type
        self.Email2DisplayName = Email2DisplayName
        self.Email3Address = Email3Address
        self.Email3Type = Email3Type
        self.Email3DisplayName = Email3DisplayName
        self.Gender = Gender
        self.GovernmentIDNumber = GovernmentIDNumber
        self.Hobby = Hobby
        self.HomeAddressPOBox = HomeAddressPOBox
        self.Initials = Initials
        self.InternetFreeBusy = InternetFreeBusy
        self.Keywords = Keywords
        self.Language = Language
        self.Location = Location
        self.ManagersName = ManagersName
        self.Mileage = Mileage
        self.Notes = Notes
        self.OfficeLocation = OfficeLocation
        self.OrganizationalIDNumber = OrganizationalIDNumber
        self.OtherAddressPOBox = OtherAddressPOBox
        self.Priority = Priority
        self.Private = Private
        self.Profession = Profession
        self.ReferredBy = ReferredBy
        self.Sensitivity = Sensitivity
        self.Spouse = Spouse
        self.User1 = User1
        self.User2 = User2
        self.User3 = User3
        self.User4 = User4
        self.WebPage = WebPage

#=====================================================================================================================
# This is (as near as I can tell) the data structure used by My Personal Address Book by Stembridge Software
#=====================================================================================================================
class NewContact(object):
    def __init__(self,
        Anniversary,
        BirthDate,
        BusAddr1,
        BusAddr2,
        BusAddr3,
        BusCity,
        BusCountry,
        BusEmail,
        BusFax,
        BusMobile,
        BusPhone,
        BusPostCode,
        BusState,
        Comment1,
        Comment2,
        Company,
        CompanyIsPrimary,
        DeathDate,
        FirstName,
        FullName,
        JobTitle,
        LastName,
        MiddleName,
        NickName,
        PerAddr1,
        PerAddr2,
        PerAddr3,
        PerCity,
        PerCountry,
        PerEmail,
        PerFax,
        PerMobile,
        PerPhone,
        PerPostCode,
        PerState,
        Sex,
        Spouse,
        Suffix,
        TimeStamp,
        Title):

        self.Anniversary = Anniversary
        self.BirthDate = BirthDate
        self.BusAddr1 = BusAddr1
        self.BusAddr2 = BusAddr2
        self.BusAddr3 = BusAddr3
        self.BusCity = BusCity
        self.BusCountry = BusCountry
        self.BusEmail = BusEmail
        self.BusFax = BusFax
        self.BusMobile = BusMobile
        self.BusPhone = BusPhone
        self.BusPostCode = BusPostCode
        self.BusState = BusState
        self.Comment1 = Comment1
        self.Comment2 = Comment2
        self.Company = Company
        self.CompanyIsPrimary = CompanyIsPrimary
        self.DeathDate = DeathDate
        self.FirstName = FirstName
        self.FullName = FullName
        self.JobTitle = JobTitle
        self.LastName = LastName
        self.MiddleName = MiddleName
        self.NickName = NickName
        self.PerAddr1 = PerAddr1
        self.PerAddr2 = PerAddr2
        self.PerAddr3 = PerAddr3
        self.PerCity = PerCity
        self.PerCountry = PerCountry
        self.PerEmail = PerEmail
        self.PerFax = PerFax
        self.PerMobile = PerMobile
        self.PerPhone = PerPhone
        self.PerPostCode = PerPostCode
        self.PerState = PerState
        self.Sex = Sex
        self.Spouse = Spouse
        self.Suffix = Suffix
        self.TimeStamp = TimeStamp
        self.Title = Title

    def __str__(self):
        return f"{self.LastName}, {self.FirstName}, {self.PerEmail}"

def loadcontacts():
    contacts    = []
    db          = sqlite3.connect('contacts.db')
    cursor      = db.cursor()
    sql_command = 'select * from contacts order by "Last Name", "First Name"'
    rows        = cursor.execute(sql_command).fetchall()
    for row in rows:
        thiscontact = Contact(*row)
        contacts.append(thiscontact)

    return contacts

def main():
    json_string = ""
    contacts = loadcontacts()
    for contact in contacts:
        newcontact = NewContact(contact.Anniversary,contact.Birthday,contact.BusinessStreet, contact.BusinessStreet2, contact.BusinessStreet3,
            contact.BusinessCity, contact.BusinessCountryRegion, None, contact.BusinessFax, contact.BusinessPhone2, contact.BusinessPhone, contact.BusinessPostalCode,
            contact.BusinessState, contact.Notes, None, contact.Company, None, None, contact.FirstName, f"{contact.LastName}, {contact.FirstName}",
            contact.JobTitle, contact.LastName, contact.MiddleName, None, contact.HomeStreet, contact.HomeStreet2, contact.HomeStreet3, contact.HomeCity,
            contact.HomeCountryRegion, contact.EmailAddress, contact.HomeFax, contact.MobilePhone, contact.HomePhone, contact.HomePostalCode,
            contact.HomeState, contact.Gender, contact.Spouse, contact.Suffix, "\\/Date(1707586353983-0500)\\/", contact.Title)
        attributes = (vars(newcontact))
        json_string = json_string + json.dumps(attributes) + ','
    prefix = '00000{"Books":[{"People":['
    suffix = '],"bookName":"My Addresses"}],"OwnerBook":null,"OwnerTimeStamp":"\\/Date(1707586353983-0500)\\/"}' #needs two backslashes so Python will print one!
    
    output = prefix + json_string[:-1] + suffix #don't need that last comma

    f = open("addrbook.txt", "w")
    f.write(output)
    f.close()

if __name__ == "__main__":
    main()