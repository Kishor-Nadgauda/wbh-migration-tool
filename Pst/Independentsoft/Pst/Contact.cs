using System;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class Contact.
    /// </summary>
    public class Contact : Item
    {
        internal Contact()
        {
        }

        #region Properties

        /// <summary>
        /// Gets the birthday.
        /// </summary>
        /// <value>The birthday.</value>
        public DateTime Birthday
        {
            get
            {
                return birthday;
            }
        }

        /// <summary>
        /// Gets the children names.
        /// </summary>
        /// <value>The children names.</value>
        public IList<string> ChildrenNames
        {
            get
            {
                return childrenNames;
            }
        }

        /// <summary>
        /// Gets the name of the assistent.
        /// </summary>
        /// <value>The name of the assistent.</value>
        public string AssistentName
        {
            get
            {
                return assistentName;
            }
        }

        /// <summary>
        /// Gets the assistent phone.
        /// </summary>
        /// <value>The assistent phone.</value>
        public string AssistentPhone
        {
            get
            {
                return assistentPhone;
            }
        }

        /// <summary>
        /// Gets the business phone.
        /// </summary>
        /// <value>The business phone.</value>
        public string BusinessPhone
        {
            get
            {
                return businessPhone;
            }
        }

        /// <summary>
        /// Gets the business fax.
        /// </summary>
        /// <value>The business fax.</value>
        public string BusinessFax
        {
            get
            {
                return businessFax;
            }
        }

        /// <summary>
        /// Gets the business home page.
        /// </summary>
        /// <value>The business home page.</value>
        public string BusinessHomePage
        {
            get
            {
                return businessHomePage;
            }
        }

        /// <summary>
        /// Gets the callback phone.
        /// </summary>
        /// <value>The callback phone.</value>
        public string CallbackPhone
        {
            get
            {
                return callbackPhone;
            }
        }

        /// <summary>
        /// Gets the car phone.
        /// </summary>
        /// <value>The car phone.</value>
        public string CarPhone
        {
            get
            {
                return carPhone;
            }
        }

        /// <summary>
        /// Gets the cellular phone.
        /// </summary>
        /// <value>The cellular phone.</value>
        public string CellularPhone
        {
            get
            {
                return cellularPhone;
            }
        }

        /// <summary>
        /// Gets the company main phone.
        /// </summary>
        /// <value>The company main phone.</value>
        public string CompanyMainPhone
        {
            get
            {
                return companyMainPhone;
            }
        }

        /// <summary>
        /// Gets the name of the company.
        /// </summary>
        /// <value>The name of the company.</value>
        public string CompanyName
        {
            get
            {
                return companyName;
            }
        }

        /// <summary>
        /// Gets the name of the computer network.
        /// </summary>
        /// <value>The name of the computer network.</value>
        public string ComputerNetworkName
        {
            get
            {
                return computerNetworkName;
            }
        }

        /// <summary>
        /// Gets the business address country.
        /// </summary>
        /// <value>The business address country.</value>
        public string BusinessAddressCountry
        {
            get
            {
                return businessAddressCountry;
            }
        }

        /// <summary>
        /// Gets the customer identifier.
        /// </summary>
        /// <value>The customer identifier.</value>
        public string CustomerId
        {
            get
            {
                return customerId;
            }
        }

        /// <summary>
        /// Gets the name of the department.
        /// </summary>
        /// <value>The name of the department.</value>
        public string DepartmentName
        {
            get
            {
                return departmentName;
            }
        }

        /// <summary>
        /// Gets the display name.
        /// </summary>
        /// <value>The display name.</value>
        public string DisplayName
        {
            get
            {
                return displayName;
            }
        }

        /// <summary>
        /// Gets the display name prefix.
        /// </summary>
        /// <value>The display name prefix.</value>
        public string DisplayNamePrefix
        {
            get
            {
                return displayNamePrefix;
            }
        }

        /// <summary>
        /// Gets the FTP site.
        /// </summary>
        /// <value>The FTP site.</value>
        public string FtpSite
        {
            get
            {
                return ftpSite;
            }
        }

        /// <summary>
        /// Gets the generation.
        /// </summary>
        /// <value>The generation.</value>
        public string Generation
        {
            get
            {
                return generation;
            }
        }

        /// <summary>
        /// Gets the name of the given.
        /// </summary>
        /// <value>The name of the given.</value>
        public string GivenName
        {
            get
            {
                return givenName;
            }
        }

        /// <summary>
        /// Gets the government identifier.
        /// </summary>
        /// <value>The government identifier.</value>
        public string GovernmentId
        {
            get
            {
                return governmentId;
            }
        }

        /// <summary>
        /// Gets the hobbies.
        /// </summary>
        /// <value>The hobbies.</value>
        public string Hobbies
        {
            get
            {
                return hobbies;
            }
        }

        /// <summary>
        /// Gets the home phone2.
        /// </summary>
        /// <value>The home phone2.</value>
        public string HomePhone2
        {
            get
            {
                return homePhone2;
            }
        }

        /// <summary>
        /// Gets the home address city.
        /// </summary>
        /// <value>The home address city.</value>
        public string HomeAddressCity
        {
            get
            {
                return homeAddressCity;
            }
        }

        /// <summary>
        /// Gets the home address country.
        /// </summary>
        /// <value>The home address country.</value>
        public string HomeAddressCountry
        {
            get
            {
                return homeAddressCountry;
            }
        }

        /// <summary>
        /// Gets the home address postal code.
        /// </summary>
        /// <value>The home address postal code.</value>
        public string HomeAddressPostalCode
        {
            get
            {
                return homeAddressPostalCode;
            }
        }

        /// <summary>
        /// Gets the home address post office box.
        /// </summary>
        /// <value>The home address post office box.</value>
        public string HomeAddressPostOfficeBox
        {
            get
            {
                return homeAddressPostOfficeBox;
            }
        }

        /// <summary>
        /// Gets the state of the home address.
        /// </summary>
        /// <value>The state of the home address.</value>
        public string HomeAddressState
        {
            get
            {
                return homeAddressState;
            }
        }

        /// <summary>
        /// Gets the home address street.
        /// </summary>
        /// <value>The home address street.</value>
        public string HomeAddressStreet
        {
            get
            {
                return homeAddressStreet;
            }
        }

        /// <summary>
        /// Gets the home fax.
        /// </summary>
        /// <value>The home fax.</value>
        public string HomeFax
        {
            get
            {
                return homeFax;
            }
        }

        /// <summary>
        /// Gets the home phone.
        /// </summary>
        /// <value>The home phone.</value>
        public string HomePhone
        {
            get
            {
                return homePhone;
            }
        }

        /// <summary>
        /// Gets the initials.
        /// </summary>
        /// <value>The initials.</value>
        public string Initials
        {
            get
            {
                return initials;
            }
        }

        /// <summary>
        /// Gets the isdn.
        /// </summary>
        /// <value>The isdn.</value>
        public string Isdn
        {
            get
            {
                return isdn;
            }
        }

        /// <summary>
        /// Gets the business address city.
        /// </summary>
        /// <value>The business address city.</value>
        public string BusinessAddressCity
        {
            get
            {
                return businessAddressCity;
            }
        }

        /// <summary>
        /// Gets the name of the manager.
        /// </summary>
        /// <value>The name of the manager.</value>
        public string ManagerName
        {
            get
            {
                return managerName;
            }
        }

        /// <summary>
        /// Gets the name of the middle.
        /// </summary>
        /// <value>The name of the middle.</value>
        public string MiddleName
        {
            get
            {
                return middleName;
            }
        }

        /// <summary>
        /// Gets the nickname.
        /// </summary>
        /// <value>The nickname.</value>
        public string Nickname
        {
            get
            {
                return nickname;
            }
        }

        /// <summary>
        /// Gets the office location.
        /// </summary>
        /// <value>The office location.</value>
        public string OfficeLocation
        {
            get
            {
                return officeLocation;
            }
        }

        /// <summary>
        /// Gets the business phone2.
        /// </summary>
        /// <value>The business phone2.</value>
        public string BusinessPhone2
        {
            get
            {
                return businessPhone2;
            }
        }

        /// <summary>
        /// Gets the other address city.
        /// </summary>
        /// <value>The other address city.</value>
        public string OtherAddressCity
        {
            get
            {
                return otherAddressCity;
            }
        }

        /// <summary>
        /// Gets the other address country.
        /// </summary>
        /// <value>The other address country.</value>
        public string OtherAddressCountry
        {
            get
            {
                return otherAddressCountry;
            }
        }

        /// <summary>
        /// Gets the other address postal code.
        /// </summary>
        /// <value>The other address postal code.</value>
        public string OtherAddressPostalCode
        {
            get
            {
                return otherAddressPostalCode;
            }
        }

        /// <summary>
        /// Gets the state of the other address.
        /// </summary>
        /// <value>The state of the other address.</value>
        public string OtherAddressState
        {
            get
            {
                return otherAddressState;
            }
        }

        /// <summary>
        /// Gets the other address street.
        /// </summary>
        /// <value>The other address street.</value>
        public string OtherAddressStreet
        {
            get
            {
                return otherAddressStreet;
            }
        }

        /// <summary>
        /// Gets the other phone.
        /// </summary>
        /// <value>The other phone.</value>
        public string OtherPhone
        {
            get
            {
                return otherPhone;
            }
        }

        /// <summary>
        /// Gets the pager.
        /// </summary>
        /// <value>The pager.</value>
        public string Pager
        {
            get
            {
                return pager;
            }
        }

        /// <summary>
        /// Gets the personal home page.
        /// </summary>
        /// <value>The personal home page.</value>
        public string PersonalHomePage
        {
            get
            {
                return personalHomePage;
            }
        }

        /// <summary>
        /// Gets the postal address.
        /// </summary>
        /// <value>The postal address.</value>
        public string PostalAddress
        {
            get
            {
                return postalAddress;
            }
        }

        /// <summary>
        /// Gets the business address postal code.
        /// </summary>
        /// <value>The business address postal code.</value>
        public string BusinessAddressPostalCode
        {
            get
            {
                return businessAddressPostalCode;
            }
        }

        /// <summary>
        /// Gets the business address post office box.
        /// </summary>
        /// <value>The business address post office box.</value>
        public string BusinessAddressPostOfficeBox
        {
            get
            {
                return businessAddressPostOfficeBox;
            }
        }

        /// <summary>
        /// Gets the state of the business address.
        /// </summary>
        /// <value>The state of the business address.</value>
        public string BusinessAddressState
        {
            get
            {
                return businessAddressState;
            }
        }

        /// <summary>
        /// Gets the business address street.
        /// </summary>
        /// <value>The business address street.</value>
        public string BusinessAddressStreet
        {
            get
            {
                return businessAddressStreet;
            }
        }

        /// <summary>
        /// Gets the primary fax.
        /// </summary>
        /// <value>The primary fax.</value>
        public string PrimaryFax
        {
            get
            {
                return primaryFax;
            }
        }

        /// <summary>
        /// Gets the primary phone.
        /// </summary>
        /// <value>The primary phone.</value>
        public string PrimaryPhone
        {
            get
            {
                return primaryPhone;
            }
        }

        /// <summary>
        /// Gets the profession.
        /// </summary>
        /// <value>The profession.</value>
        public string Profession
        {
            get
            {
                return profession;
            }
        }

        /// <summary>
        /// Gets the radio phone.
        /// </summary>
        /// <value>The radio phone.</value>
        public string RadioPhone
        {
            get
            {
                return radioPhone;
            }
        }

        /// <summary>
        /// Gets the name of the spouse.
        /// </summary>
        /// <value>The name of the spouse.</value>
        public string SpouseName
        {
            get
            {
                return spouseName;
            }
        }

        /// <summary>
        /// Gets the surname.
        /// </summary>
        /// <value>The surname.</value>
        public string Surname
        {
            get
            {
                return surname;
            }
        }

        /// <summary>
        /// Gets the telex.
        /// </summary>
        /// <value>The telex.</value>
        public string Telex
        {
            get
            {
                return telex;
            }
        }

        /// <summary>
        /// Gets the title.
        /// </summary>
        /// <value>The title.</value>
        public string Title
        {
            get
            {
                return title;
            }
        }

        /// <summary>
        /// Gets the tty TDD phone.
        /// </summary>
        /// <value>The tty TDD phone.</value>
        public string TtyTddPhone
        {
            get
            {
                return ttyTddPhone;
            }
        }

        /// <summary>
        /// Gets the wedding anniversary.
        /// </summary>
        /// <value>The wedding anniversary.</value>
        public DateTime WeddingAnniversary
        {
            get
            {
                return weddingAnniversary;
            }
        }

        /// <summary>
        /// Gets the gender.
        /// </summary>
        /// <value>The gender.</value>
        public Gender Gender
        {
            get
            {
                return gender;
            }
        }

        /// <summary>
        /// Gets the selected mailing address.
        /// </summary>
        /// <value>The selected mailing address.</value>
        public SelectedMailingAddress SelectedMailingAddress
        {
            get
            {
                return selectedMailingAddress;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance has picture.
        /// </summary>
        /// <value><c>true</c> if this instance has picture; otherwise, <c>false</c>.</value>
        public bool HasPicture
        {
            get
            {
                return contactHasPicture;
            }
        }

        /// <summary>
        /// Gets the file as.
        /// </summary>
        /// <value>The file as.</value>
        public string FileAs
        {
            get
            {
                return fileAs;
            }
        }

        /// <summary>
        /// Gets the instant messenger address.
        /// </summary>
        /// <value>The instant messenger address.</value>
        public string InstantMessengerAddress
        {
            get
            {
                return instantMessengerAddress;
            }
        }

        /// <summary>
        /// Gets the internet free busy address.
        /// </summary>
        /// <value>The internet free busy address.</value>
        public string InternetFreeBusyAddress
        {
            get
            {
                return internetFreeBusyAddress;
            }
        }

        /// <summary>
        /// Gets the business address.
        /// </summary>
        /// <value>The business address.</value>
        public string BusinessAddress
        {
            get
            {
                return businessAddress;
            }
        }

        /// <summary>
        /// Gets the home address.
        /// </summary>
        /// <value>The home address.</value>
        public string HomeAddress
        {
            get
            {
                return homeAddress;
            }
        }

        /// <summary>
        /// Gets the other address.
        /// </summary>
        /// <value>The other address.</value>
        public string OtherAddress
        {
            get
            {
                return otherAddress;
            }
        }

        /// <summary>
        /// Gets the email1 address.
        /// </summary>
        /// <value>The email1 address.</value>
        public string Email1Address
        {
            get
            {
                return email1Address;
            }
        }

        /// <summary>
        /// Gets the email2 address.
        /// </summary>
        /// <value>The email2 address.</value>
        public string Email2Address
        {
            get
            {
                return email2Address;
            }
        }

        /// <summary>
        /// Gets the email3 address.
        /// </summary>
        /// <value>The email3 address.</value>
        public string Email3Address
        {
            get
            {
                return email3Address;
            }
        }

        /// <summary>
        /// Gets the display name of the email1.
        /// </summary>
        /// <value>The display name of the email1.</value>
        public string Email1DisplayName
        {
            get
            {
                return email1DisplayName;
            }
        }

        /// <summary>
        /// Gets the display name of the email2.
        /// </summary>
        /// <value>The display name of the email2.</value>
        public string Email2DisplayName
        {
            get
            {
                return email2DisplayName;
            }
        }

        /// <summary>
        /// Gets the display name of the email3.
        /// </summary>
        /// <value>The display name of the email3.</value>
        public string Email3DisplayName
        {
            get
            {
                return email3DisplayName;
            }
        }

        /// <summary>
        /// Gets the email1 display as.
        /// </summary>
        /// <value>The email1 display as.</value>
        public string Email1DisplayAs
        {
            get
            {
                return email1DisplayAs;
            }
        }

        /// <summary>
        /// Gets the email2 display as.
        /// </summary>
        /// <value>The email2 display as.</value>
        public string Email2DisplayAs
        {
            get
            {
                return email2DisplayAs;
            }
        }

        /// <summary>
        /// Gets the email3 display as.
        /// </summary>
        /// <value>The email3 display as.</value>
        public string Email3DisplayAs
        {
            get
            {
                return email3DisplayAs;
            }
        }

        /// <summary>
        /// Gets the type of the email1.
        /// </summary>
        /// <value>The type of the email1.</value>
        public string Email1Type
        {
            get
            {
                return email1Type;
            }
        }

        /// <summary>
        /// Gets the type of the email2.
        /// </summary>
        /// <value>The type of the email2.</value>
        public string Email2Type
        {
            get
            {
                return email2Type;
            }
        }

        /// <summary>
        /// Gets the type of the email3.
        /// </summary>
        /// <value>The type of the email3.</value>
        public string Email3Type
        {
            get
            {
                return email3Type;
            }
        }

        /// <summary>
        /// Gets the email1 entry identifier.
        /// </summary>
        /// <value>The email1 entry identifier.</value>
        public byte[] Email1EntryId
        {
            get
            {
                return email1EntryId;
            }
        }

        /// <summary>
        /// Gets the email2 entry identifier.
        /// </summary>
        /// <value>The email2 entry identifier.</value>
        public byte[] Email2EntryId
        {
            get
            {
                return email2EntryId;
            }
        }

        /// <summary>
        /// Gets the email3 entry identifier.
        /// </summary>
        /// <value>The email3 entry identifier.</value>
        public byte[] Email3EntryId
        {
            get
            {
                return email3EntryId;
            }
        }

        #endregion
    }
}
