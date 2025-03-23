import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import styles from './EnquiryWebPart.module.scss';
import * as strings from 'EnquiryWebPartStrings';

export interface IEnquiryWebPartProps {
  formTitle: string;
  submitButtonText: string;
  thankYouMessage: string;
  faqPageUrl: string;
  submissionListName: string;
  documentLibraryName: string;
  notificationEmail: string;
  adminGroupName: string;
}

// Country list for dropdown
const COUNTRIES = [
  "Afghanistan", "Albania", "Algeria", "Andorra", "Angola", "Antigua and Barbuda", "Argentina", "Armenia", "Australia", "Austria", 
  "Azerbaijan", "Bahamas", "Bahrain", "Bangladesh", "Barbados", "Belarus", "Belgium", "Belize", "Benin", "Bhutan", "Bolivia", 
  "Bosnia and Herzegovina", "Botswana", "Brazil", "Brunei", "Bulgaria", "Burkina Faso", "Burundi", "Cabo Verde", "Cambodia", 
  "Cameroon", "Canada", "Central African Republic", "Chad", "Chile", "China", "Colombia", "Comoros", "Congo", "Costa Rica", 
  "Croatia", "Cuba", "Cyprus", "Czech Republic", "Denmark", "Djibouti", "Dominica", "Dominican Republic", "Ecuador", "Egypt", 
  "El Salvador", "Equatorial Guinea", "Eritrea", "Estonia", "Eswatini", "Ethiopia", "Fiji", "Finland", "France", "Gabon", "Gambia", 
  "Georgia", "Germany", "Ghana", "Greece", "Grenada", "Guatemala", "Guinea", "Guinea-Bissau", "Guyana", "Haiti", "Honduras", 
  "Hungary", "Iceland", "India", "Indonesia", "Iran", "Iraq", "Ireland", "Israel", "Italy", "Jamaica", "Japan", "Jordan", "Kazakhstan", 
  "Kenya", "Kiribati", "Korea, North", "Korea, South", "Kosovo", "Kuwait", "Kyrgyzstan", "Laos", "Latvia", "Lebanon", "Lesotho", 
  "Liberia", "Libya", "Liechtenstein", "Lithuania", "Luxembourg", "Madagascar", "Malawi", "Malaysia", "Maldives", "Mali", "Malta", 
  "Marshall Islands", "Mauritania", "Mauritius", "Mexico", "Micronesia", "Moldova", "Monaco", "Mongolia", "Montenegro", "Morocco", 
  "Mozambique", "Myanmar", "Namibia", "Nauru", "Nepal", "Netherlands", "New Zealand", "Nicaragua", "Niger", "Nigeria", "North Macedonia", 
  "Norway", "Oman", "Pakistan", "Palau", "Palestine", "Panama", "Papua New Guinea", "Paraguay", "Peru", "Philippines", "Poland", 
  "Portugal", "Qatar", "Romania", "Russia", "Rwanda", "Saint Kitts and Nevis", "Saint Lucia", "Saint Vincent and the Grenadines", 
  "Samoa", "San Marino", "Sao Tome and Principe", "Saudi Arabia", "Senegal", "Serbia", "Seychelles", "Sierra Leone", "Singapore", 
  "Slovakia", "Slovenia", "Solomon Islands", "Somalia", "South Africa", "South Sudan", "Spain", "Sri Lanka", "Sudan", "Suriname", 
  "Sweden", "Switzerland", "Syria", "Taiwan", "Tajikistan", "Tanzania", "Thailand", "Timor-Leste", "Togo", "Tonga", "Trinidad and Tobago", 
  "Tunisia", "Turkey", "Turkmenistan", "Tuvalu", "Uganda", "Ukraine", "United Arab Emirates", "United Kingdom", "United States", 
  "Uruguay", "Uzbekistan", "Vanuatu", "Vatican City", "Venezuela", "Vietnam", "Yemen", "Zambia", "Zimbabwe", "Global Presence"
];

// Constants for the form
const REGULATORS = [
  "Financial Conduct Authority (FCA)",
  "Prudential Regulation Authority (PRA)",
  "Financial Policy Committee (FPC)",
  "Payment Systems Regulator (PSR)",
  "Information Commissioner's Office (ICO)",
  "Advertising Standards Authority (ASA)",
  "Competition and Markets Authority (CMA)",
  "Other"
];

const CATEGORIES = {
  'banking': [
    'Digital/Neo-bank',
    'Business/Corporate Banking',
    'Retail Banking',
    'Private Banking',
    'Investment Banking',
    'Commercial Banking',
    'Core Banking Technology',
    'Other'
  ],
  'payments': [
    'Payment Institution',
    'E-money Institution',
    'Money Transfer',
    'Acquiring Services',
    'Card Issuance',
    'Digital Wallet',
    'Open Banking Solutions',
    'Other'
  ],
  'insurance': [
    'Life Insurance',
    'General Insurance',
    'Health Insurance',
    'Parametric Insurance',
    'Insurance Brokering',
    'Other'
  ],
  'crypto': [
    'Exchange',
    'Wallet Provider',
    'Stablecoin',
    'Decentralized Finance (DeFi)',
    'Crypto Lending/Borrowing',
    'NFT Marketplace',
    'Other'
  ],
  'investment': [
    'Wealth Management',
    'Robo-advisory',
    'Crowd-funding',
    'P2P Lending',
    'Asset Management',
    'Other'
  ],
  'credit': [
    'Lending',
    'Buy Now Pay Later',
    'Credit Scoring',
    'Credit Reference',
    'Debt Management',
    'Other'
  ],
  'notOperational': [
    'Pre-launch Product/Service',
    'Research and Development',
    'Early-stage Product/Service',
    'Other'
  ]
};

export default class EnquiryWebPart extends BaseClientSideWebPart<IEnquiryWebPartProps> {

  private currentStep: number = 1;
  private formData: any = {
    // Basic Information
    fullName: '',
    organisationName: '',
    contactNumber: '',
    emailAddress: '',
    websiteAddress: '',
    operationLocation: '',
    countriesOfOperation: [],
    operationLength: '',
    
    // Industry Information
    primaryBusinessAreas: '',
    productServiceCategory: '',
    otherProductServiceCategory: '',
    operationalStatus: null,
    regulatoryStatus: null,
    regulators: [],
    otherRegulator: '',
    
    // Enquiry Details
    productServiceDescription: '',
    questions: [''],
    additionalInformation: '',
    faqConfirmation: null,
    consentConfirmation: false,
    
    // Attachments
    files: []
  };

  // Number of questions currently shown (starts with 1)
  private questionCount: number = 1;
  
  // File upload references
  private fileUploadElement: HTMLInputElement = null;
  private uploadedFiles: File[] = [];
  
  // Validation flag
  private validateAttempted: boolean = false;

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.enquiry }">
        <div class="${ styles.container }">
          <div class="${ styles.header }">
            <h2 class="${ styles.title }">Regulatory Guidance Unit Enquiry Form</h2>
            <div class="${ styles.introText }">
              <p>The Regulatory Guidance Unit provides informal, non-binding steers to persons seeking direction and clarity in navigating aspects of the FinTech regulatory landscape. It relies on the expertise of representatives from across participating regulators within the IFWG to ensure that guidance is holistic, inclusive and well considered. The functions of the Regulatory Guidance Unit include the following:</p>
              <ul>
                <li>To provide innovators with efficient and effective access to regulatory expertise potentially reducing their time needed to resolve regulatory concerns, increasing their speed to market and lowering their legal fees.</li>
                <li>To impart guidance and insight to entities seeking to operate and innovate in the market.</li>
                <li>To improve regulatory compliance through clear articulation of regulatory frameworks and reduce regulatory arbitrage.</li>
                <li>To enhance regulator understanding of innovation in the market.</li>
              </ul>
              <p>All information provided below will be kept confidential in accordance with the Protection of Personal Information Act, 2013, and the IFWG Privacy Policy.</p>
              <p>Please ensure you have consulted the <a href="${escape(this.properties.faqPageUrl || '#')}" target="_blank">FAQs page</a> before submitting an enquiry. <strong>All enquiries that relate to, and are addressed by, the FAQs will be referred to the FAQs page.</strong></p>
              <p><em>* Mandatory fields</em></p>
            </div>
            <div class="${ styles.progressContainer }">
              <div class="${ styles.progressBar }">
                <div class="${ styles.progress }" style="width: ${(this.currentStep - 1) * 33.33}%"></div>
              </div>
              <div class="${ styles.steps }">
                <div class="${ styles.step } ${this.currentStep >= 1 ? styles.active : ''} ${this.currentStep === 1 ? styles.current : ''}">
                  <div class="${ styles.stepNumber }">1</div>
                  <div class="${ styles.stepLabel }">Basic Information</div>
                </div>
                <div class="${ styles.step } ${this.currentStep >= 2 ? styles.active : ''} ${this.currentStep === 2 ? styles.current : ''}">
                  <div class="${ styles.stepNumber }">2</div>
                  <div class="${ styles.stepLabel }">Industry Information</div>
                </div>
                <div class="${ styles.step } ${this.currentStep >= 3 ? styles.active : ''} ${this.currentStep === 3 ? styles.current : ''}">
                  <div class="${ styles.stepNumber }">3</div>
                  <div class="${ styles.stepLabel }">Your Enquiry</div>
                </div>
              </div>
            </div>
          </div>
          
          <div class="${ styles.formContainer }">
            ${this.renderCurrentStep()}
          </div>
        </div>
      </div>`;

    this.setButtonHandlers();
  }

  private renderCurrentStep(): string {
    switch (this.currentStep) {
      case 1:
        return this.renderBasicInfoStep();
      case 2:
        return this.renderIndustryStep();
      case 3:
        return this.renderInquiryStep();
      case 4:
        return this.renderThankYouStep();
      default:
        return this.renderBasicInfoStep();
    }
  }

  private renderBasicInfoStep(): string {
    return `
      <div class="${ styles.formStep } ${styles.fadeIn}">
        <h3 class="${ styles.stepTitle }">Section A: Basic Information</h3>
        
        <div class="${ styles.formField }">
          <label for="fullName">Full Name <span class="${ styles.required }">*</span></label>
          <input type="text" id="fullName" class="${ styles.textField }${this.formData.fullName === '' && this.validateAttempted ? ' ' + styles.error : ''}" value="${this.formData.fullName || ''}">
          ${this.formData.fullName === '' && this.validateAttempted ? `<div class="${styles.errorText}">Full Name is required</div>` : ''}
        </div>
        
        <div class="${ styles.formField }">
          <label for="organisationName">Organisation Name <span class="${ styles.required }">*</span></label>
          <input type="text" id="organisationName" class="${ styles.textField }${this.formData.organisationName === '' && this.validateAttempted ? ' ' + styles.error : ''}" value="${this.formData.organisationName || ''}">
          ${this.formData.organisationName === '' && this.validateAttempted ? `<div class="${styles.errorText}">Organisation Name is required</div>` : ''}
        </div>
        
        <div class="${ styles.formField }">
          <label for="contactNumber">Contact Number</label>
          <input type="tel" id="contactNumber" class="${ styles.textField }" value="${this.formData.contactNumber || ''}">
        </div>
        
        <div class="${ styles.formField }">
          <label for="emailAddress">Email Address <span class="${ styles.required }">*</span></label>
          <input type="email" id="emailAddress" class="${ styles.textField }${(this.formData.emailAddress === '' || !this.isValidEmail(this.formData.emailAddress)) && this.validateAttempted ? ' ' + styles.error : ''}" value="${this.formData.emailAddress || ''}">
          ${this.formData.emailAddress === '' && this.validateAttempted ? `<div class="${styles.errorText}">Email Address is required</div>` : ''}
          ${this.formData.emailAddress !== '' && !this.isValidEmail(this.formData.emailAddress) && this.validateAttempted ? `<div class="${styles.errorText}">Please enter a valid email address</div>` : ''}
        </div>
        
        <div class="${ styles.formField }">
          <label for="websiteAddress">Website Address</label>
          <input type="url" id="websiteAddress" class="${ styles.textField }" value="${this.formData.websiteAddress || ''}">
        </div>
        
        <div class="${ styles.formField }">
          <label for="operationLocation">Where do you operate/intend to operate? <span class="${ styles.required }">*</span></label>
          <select id="operationLocation" class="${ styles.selectField }${this.formData.operationLocation === '' && this.validateAttempted ? ' ' + styles.error : ''}">
            <option value="" ${!this.formData.operationLocation ? 'selected' : ''}>Please select</option>
            <option value="South Africa only" ${this.formData.operationLocation === 'South Africa only' ? 'selected' : ''}>I operate in South Africa only</option>
            <option value="South Africa + other countries" ${this.formData.operationLocation === 'South Africa + other countries' ? 'selected' : ''}>I currently only operate in South Africa but intend to operate in other countries</option>
            <option value="Not operational - South Africa only" ${this.formData.operationLocation === 'Not operational - South Africa only' ? 'selected' : ''}>I am not yet operational but intend to operate only in South Africa</option>
            <option value="Not operational - South Africa + other countries" ${this.formData.operationLocation === 'Not operational - South Africa + other countries' ? 'selected' : ''}>I am not yet operational, but intend to operate in South Africa and other countries</option>
            <option value="Outside South Africa - plan to operate in South Africa" ${this.formData.operationLocation === 'Outside South Africa - plan to operate in South Africa' ? 'selected' : ''}>I operate outside of South Africa only, but intend to operate in South Africa</option>
          </select>
          ${this.formData.operationLocation === '' && this.validateAttempted ? `<div class="${styles.errorText}">Please select your operation location</div>` : ''}
        </div>
        
        <div class="${ styles.formField }">
          <label>Please specify the countries in which you operate/intend operating? <span class="${ styles.required }">*</span></label>
          <div class="${ styles.info }">i
            <div class="${ styles.tooltip }">If you operate in more than 5 countries, please check the box marked "global presence"</div>
          </div>
          <div class="${ styles.countrySelector }">
            <div class="${ styles.searchBox }">
              <input type="text" id="countrySearch" class="${ styles.textField }" placeholder="Search countries..." />
            </div>
            <div class="${ styles.countryList }">
              ${COUNTRIES.map(country => `
                <div class="${ styles.countryItem }">
                  <input type="checkbox" id="country-${country.replace(/\s+/g, '-').toLowerCase()}" 
                    class="country-checkbox" value="${country}" 
                    ${this.formData.countriesOfOperation.includes(country) ? 'checked' : ''} />
                  <label for="country-${country.replace(/\s+/g, '-').toLowerCase()}">${country}</label>
                </div>
              `).join('')}
            </div>
            <div class="${ styles.selectedCountries }${this.formData.countriesOfOperation.length === 0 && this.validateAttempted ? ' ' + styles.error : ''}">
              ${this.formData.countriesOfOperation.length > 0 ? 
                this.formData.countriesOfOperation.map(country => `
                  <div class="${ styles.countryTag }" data-country="${country}">
                    ${country} <span class="${ styles.removeCountry }" data-country="${country}">×</span>
                  </div>
                `).join('') : 
                '<div>No countries selected</div>'
              }
            </div>
            ${this.formData.countriesOfOperation.length === 0 && this.validateAttempted ? `<div class="${styles.errorText}">Please select at least one country</div>` : ''}
          </div>
        </div>
        
        <div class="${ styles.formField }">
          <label for="operationLength">For how long has the organisation been in operation? <span class="${ styles.required }">*</span></label>
          <select id="operationLength" class="${ styles.selectField }${this.formData.operationLength === '' && this.validateAttempted ? ' ' + styles.error : ''}">
            <option value="" ${!this.formData.operationLength ? 'selected' : ''}>Please select</option>
            <option value="Not operational" ${this.formData.operationLength === 'Not operational' ? 'selected' : ''}>Not operational</option>
            <option value="<1 year" ${this.formData.operationLength === '<1 year' ? 'selected' : ''}>Less than 1 year</option>
            <option value="1-5 years" ${this.formData.operationLength === '1-5 years' ? 'selected' : ''}>1 – 5 years</option>
            <option value="5-10 years" ${this.formData.operationLength === '5-10 years' ? 'selected' : ''}>5 – 10 years</option>
            <option value=">10 years" ${this.formData.operationLength === '>10 years' ? 'selected' : ''}>More than 10 years</option>
          </select>
          ${this.formData.operationLength === '' && this.validateAttempted ? `<div class="${styles.errorText}">Please select how long your organisation has been in operation</div>` : ''}
        </div>
        
        <div class="${ styles.formActions }">
          <button type="button" class="${ styles.button } ${styles.nextButton}" id="nextToStep2">Next</button>
        </div>
      </div>
    `;
  }

  private renderIndustryStep(): string {
    return `
      <div class="${ styles.formStep } ${styles.fadeIn}">
        <h3 class="${ styles.stepTitle }">Section B: Industry Information</h3>
        <div class="${ styles.sectionNote }">
          For the next section you may select more than one option if it applies to your organisation
        </div>
        
        <div class="${ styles.formField }">
          <label for="primaryBusinessAreas">What is the primary area of business in which you currently operate? <span class="${ styles.required }">*</span></label>
          <select id="primaryBusinessAreas" class="${ styles.selectField }${this.formData.primaryBusinessAreas === '' && this.validateAttempted ? ' ' + styles.error : ''}">
            <option value="" ${!this.formData.primaryBusinessAreas ? 'selected' : ''}>Please select</option>
            <option value="banking" ${this.formData.primaryBusinessAreas === 'banking' ? 'selected' : ''}>Banking</option>
            <option value="payments" ${this.formData.primaryBusinessAreas === 'payments' ? 'selected' : ''}>Payments</option>
            <option value="insurance" ${this.formData.primaryBusinessAreas === 'insurance' ? 'selected' : ''}>Insurance</option>
            <option value="crypto" ${this.formData.primaryBusinessAreas === 'crypto' ? 'selected' : ''}>Crypto</option>
            <option value="investment" ${this.formData.primaryBusinessAreas === 'investment' ? 'selected' : ''}>Investment</option>
            <option value="credit" ${this.formData.primaryBusinessAreas === 'credit' ? 'selected' : ''}>Credit</option>
            <option value="notOperational" ${this.formData.primaryBusinessAreas === 'notOperational' ? 'selected' : ''}>Not operational</option>
          </select>
          ${this.formData.primaryBusinessAreas === '' && this.validateAttempted ? `<div class="${styles.errorText}">Please select your primary business area</div>` : ''}
        </div>
        
        <div class="${ styles.formField }">
          <label for="productServiceCategory">Which product/service category does your enquiry relate to? <span class="${ styles.required }">*</span></label>
          <select id="productServiceCategory" class="${ styles.selectField }${this.formData.productServiceCategory === '' && this.validateAttempted ? ' ' + styles.error : ''}">
            <option value="" ${!this.formData.productServiceCategory ? 'selected' : ''}>Please select</option>
            ${this.renderProductServiceCategoryOptions()}
          </select>
          ${this.formData.productServiceCategory === '' && this.validateAttempted ? `<div class="${styles.errorText}">Please select a product/service category</div>` : ''}
        </div>
        
        ${this.formData.productServiceCategory === 'Other' ? `
        <div class="${ styles.formField }">
          <label for="otherProductServiceCategory">Other product/service category <span class="${ styles.required }">*</span></label>
          <input type="text" id="otherProductServiceCategory" class="${ styles.textField }${this.formData.otherProductServiceCategory === '' && this.validateAttempted ? ' ' + styles.error : ''}" value="${this.formData.otherProductServiceCategory || ''}">
          ${this.formData.otherProductServiceCategory === '' && this.validateAttempted ? `<div class="${styles.errorText}">Please specify the other product/service category</div>` : ''}
        </div>
        ` : ''}
        
        <div class="${ styles.formField }">
          <label>Is the product/service to which the enquiry relates operational in the market? <span class="${ styles.required }">*</span></label>
          <div class="${ styles.radioGroup }${this.formData.operationalStatus === null && this.validateAttempted ? ' ' + styles.error : ''}">
            <div class="${ styles.radioItem }">
              <input type="radio" id="operationalStatusYes" name="operationalStatus" value="yes" ${this.formData.operationalStatus === true ? 'checked' : ''}>
              <label for="operationalStatusYes">Yes</label>
            </div>
            <div class="${ styles.radioItem }">
              <input type="radio" id="operationalStatusNo" name="operationalStatus" value="no" ${this.formData.operationalStatus === false ? 'checked' : ''}>
              <label for="operationalStatusNo">No</label>
            </div>
          </div>
          ${this.formData.operationalStatus === null && this.validateAttempted ? `<div class="${styles.errorText}">Please indicate if the product/service is operational</div>` : ''}
        </div>
        
        <div class="${ styles.formField }">
          <label>Are you licensed/authorised/registered by regulatory authorities to undertake your primary business? <span class="${ styles.required }">*</span></label>
          <div class="${ styles.radioGroup }${this.formData.regulatoryStatus === null && this.validateAttempted ? ' ' + styles.error : ''}">
            <div class="${ styles.radioItem }">
              <input type="radio" id="regulatoryStatusYes" name="regulatoryStatus" value="yes" ${this.formData.regulatoryStatus === true ? 'checked' : ''}>
              <label for="regulatoryStatusYes">Yes</label>
            </div>
            <div class="${ styles.radioItem }">
              <input type="radio" id="regulatoryStatusNo" name="regulatoryStatus" value="no" ${this.formData.regulatoryStatus === false ? 'checked' : ''}>
              <label for="regulatoryStatusNo">No</label>
            </div>
          </div>
          ${this.formData.regulatoryStatus === null && this.validateAttempted ? `<div class="${styles.errorText}">Please indicate your regulatory status</div>` : ''}
        </div>
        
        ${this.formData.regulatoryStatus === true ? `
        <div class="${ styles.formField }">
          <label>Select regulators <span class="${ styles.required }">*</span></label>
          <div class="${ styles.regulatorSelector }${this.formData.regulators.length === 0 && this.validateAttempted ? ' ' + styles.error : ''}">
            ${REGULATORS.map(regulator => `
              <div class="${ styles.checkboxItem }">
                <input type="checkbox" id="regulator-${regulator.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase()}" 
                  class="regulator-checkbox" value="${regulator}" 
                  ${this.formData.regulators.indexOf(regulator) !== -1 ? 'checked' : ''} />
                <label for="regulator-${regulator.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase()}">${regulator}</label>
              </div>
            `).join('')}
          </div>
          ${this.formData.regulators.length === 0 && this.validateAttempted ? `<div class="${styles.errorText}">Please select at least one regulator</div>` : ''}
        </div>
        ` : ''}
        
        ${this.formData.regulators.indexOf('Other') !== -1 ? `
        <div class="${ styles.formField }">
          <label for="otherRegulator">Please specify other regulator <span class="${ styles.required }">*</span></label>
          <input type="text" id="otherRegulator" class="${ styles.textField }${this.formData.otherRegulator === '' && this.validateAttempted ? ' ' + styles.error : ''}" value="${this.formData.otherRegulator || ''}">
          ${this.formData.otherRegulator === '' && this.validateAttempted ? `<div class="${styles.errorText}">Please specify the other regulator</div>` : ''}
        </div>
        ` : ''}
        
        <div class="${ styles.formActions }">
          <button type="button" class="${ styles.button } ${styles.backButton}" id="backToStep1">Back</button>
          <button type="button" class="${ styles.button } ${styles.nextButton}" id="nextToStep3">Next</button>
        </div>
      </div>
    `;
  }
  
  private renderProductServiceCategoryOptions(): string {
    const businessArea = this.formData.primaryBusinessAreas;
    if (!businessArea || !CATEGORIES[businessArea]) {
      return '';
    }
    
    let options = '';
    const categories = CATEGORIES[businessArea];
    for (let i = 0; i < categories.length; i++) {
      const category = categories[i];
      options += `<option value="${category}" ${this.formData.productServiceCategory === category ? 'selected' : ''}>${category}</option>`;
    }
    
    return options;
  }

  private renderInquiryStep(): string {
    return `
      <div class="${ styles.formStep } ${styles.fadeIn}">
        <h3 class="${ styles.stepTitle }">Section C: Inquiry Details</h3>
        <div class="${ styles.sectionNote }">
          Please provide additional details about your product/service and specific questions you have
        </div>
        
        <div class="${ styles.formField }">
          <label for="productServiceDescription">Please describe your product/service and its key features <span class="${ styles.required }">*</span></label>
          <textarea id="productServiceDescription" class="${ styles.textareaField }${this.formData.productServiceDescription === '' && this.validateAttempted ? ' ' + styles.error : ''}" rows="5">${this.formData.productServiceDescription || ''}</textarea>
          ${this.formData.productServiceDescription === '' && this.validateAttempted ? `<div class="${styles.errorText}">Please provide a description of your product/service</div>` : ''}
        </div>
        
        <div class="${ styles.formField }">
          <label>What are your specific questions for the Regulatory Guidance Unit? <span class="${ styles.required }">*</span></label>
          <div class="${ styles.questionsContainer }${this.formData.questions[0] === '' && this.validateAttempted ? ' ' + styles.error : ''}">
            ${this.renderQuestions()}
          </div>
          ${this.formData.questions[0] === '' && this.validateAttempted ? `<div class="${styles.errorText}">Please provide at least one question</div>` : ''}
          <button type="button" class="${ styles.button } ${styles.addButton}" id="addQuestionBtn">Add Another Question</button>
        </div>
        
        <div class="${ styles.formField }">
          <label for="additionalInformation">Any additional information related to your inquiry?</label>
          <textarea id="additionalInformation" class="${ styles.textareaField }" rows="3">${this.formData.additionalInformation || ''}</textarea>
        </div>
        
        <div class="${ styles.formField }">
          <label>Have you checked our FAQs for this information? <span class="${ styles.required }">*</span></label>
          <div class="${ styles.radioGroup }${this.formData.faqConfirmation === null && this.validateAttempted ? ' ' + styles.error : ''}">
            <div class="${ styles.radioItem }">
              <input type="radio" id="faqConfirmationYes" name="faqConfirmation" value="yes" ${this.formData.faqConfirmation === true ? 'checked' : ''}>
              <label for="faqConfirmationYes">Yes</label>
            </div>
            <div class="${ styles.radioItem }">
              <input type="radio" id="faqConfirmationNo" name="faqConfirmation" value="no" ${this.formData.faqConfirmation === false ? 'checked' : ''}>
              <label for="faqConfirmationNo">No</label>
            </div>
          </div>
          ${this.formData.faqConfirmation === null && this.validateAttempted ? `<div class="${styles.errorText}">Please indicate if you've checked our FAQs</div>` : ''}
          <div class="${ styles.faqLink }">
            <a href="${this.properties.faqPageUrl || '#'}" target="_blank">View FAQs</a>
          </div>
        </div>
        
        <div class="${ styles.formField }">
          <label>Supporting Files</label>
          <div class="${ styles.fileUploadContainer }">
            <input type="file" id="fileUpload" multiple class="${ styles.fileInput }" />
            <button type="button" class="${ styles.button } ${styles.uploadButton}" id="uploadBtn">Upload Files</button>
          </div>
          <div class="${ styles.uploadedFiles }" id="uploadedFilesContainer">
            ${this.renderUploadedFiles()}
          </div>
          <div class="${ styles.uploadNote }">
            Maximum 5 files allowed. Accepted formats: PDF, DOC, DOCX, XLS, XLSX, PNG, JPG (max 10MB each)
          </div>
        </div>
        
        <div class="${ styles.formField }">
          <div class="${ styles.checkbox }${!this.formData.consentConfirmation && this.validateAttempted ? ' ' + styles.error : ''}">
            <input type="checkbox" id="consentCheckbox" ${this.formData.consentConfirmation ? 'checked' : ''}>
            <label for="consentCheckbox">I consent to my information being processed in accordance with the privacy policy <span class="${ styles.required }">*</span></label>
          </div>
          ${!this.formData.consentConfirmation && this.validateAttempted ? `<div class="${styles.errorText}">You must provide consent to submit this form</div>` : ''}
        </div>
        
        <div class="${ styles.formActions }">
          <button type="button" class="${ styles.button } ${styles.backButton}" id="backToStep2">Back</button>
          <button type="button" class="${ styles.button } ${styles.submitButton}" id="submitBtn">${escape(this.properties.submitButtonText || 'Submit')}</button>
        </div>
      </div>
    `;
  }
  
  private renderQuestions(): string {
    let questionsHtml = '';
    
    for (let i = 0; i < this.formData.questions.length; i++) {
      questionsHtml += `
        <div class="${ styles.questionItem }">
          <textarea class="${ styles.textareaField } question-input" rows="2">${this.formData.questions[i] || ''}</textarea>
          ${i > 0 ? `<button type="button" class="${ styles.removeButton } remove-question-btn" data-index="${i}">Remove</button>` : ''}
        </div>
      `;
    }
    
    return questionsHtml;
  }
  
  private renderUploadedFiles(): string {
    if (!this.formData.files || this.formData.files.length === 0) {
      return '<div class="no-files">No files uploaded</div>';
    }
    
    let filesHtml = '';
    
    for (let i = 0; i < this.formData.files.length; i++) {
      const file = this.formData.files[i];
      filesHtml += `
        <div class="${ styles.fileItem }">
          <span class="${ styles.fileName }">${file.name}</span>
          <button type="button" class="${ styles.removeButton } remove-file-btn" data-index="${i}">Remove</button>
        </div>
      `;
    }
    
    return filesHtml;
  }

  private renderThankYouStep(): string {
    return `
      <div class="${ styles.formStep } ${styles.fadeIn} ${styles.thankYouStep}">
        <div class="${ styles.thankYouIcon }">
          <i class="${ styles.checkmark }">✓</i>
        </div>
        <h3 class="${ styles.thankYouTitle }">Thank You!</h3>
        <p class="${ styles.thankYouMessage }">${escape(this.properties.thankYouMessage || 'Your inquiry has been submitted successfully. We will contact you soon.')}</p>
        <div class="${ styles.formActions }">
          <button type="button" class="${ styles.button } ${styles.newInquiryButton}" id="newInquiryBtn">Submit Another Inquiry</button>
        </div>
      </div>
    `;
  }

  private setButtonHandlers(): void {
    console.log('Setting up button handlers for step:', this.currentStep);
    
    // Navigation buttons
    const nextToStep2Button = this.domElement.querySelector('#nextToStep2');
    if (nextToStep2Button) {
      console.log('Found Next to Step 2 button');
      nextToStep2Button.addEventListener('click', () => {
        console.log('Next to Step 2 clicked');
        if (this.validateStep1()) {
          this.saveStep1Data();
          this.currentStep = 2;
          this.validateAttempted = false;
          this.render();
        }
      });
    } else {
      console.log('Next to Step 2 button not found');
    }
    
    const backToStep1Button = this.domElement.querySelector('#backToStep1');
    if (backToStep1Button) {
      backToStep1Button.addEventListener('click', () => {
        console.log('Back to Step 1 clicked');
        this.currentStep = 1;
        this.validateAttempted = false;
        this.render();
      });
    }
    
    const nextToStep3Button = this.domElement.querySelector('#nextToStep3');
    if (nextToStep3Button) {
      nextToStep3Button.addEventListener('click', () => {
        console.log('Next to Step 3 clicked');
        if (this.validateStep2()) {
          this.saveStep2Data();
          this.currentStep = 3;
          this.validateAttempted = false;
          this.render();
        }
      });
    }
    
    const backToStep2Button = this.domElement.querySelector('#backToStep2');
    if (backToStep2Button) {
      backToStep2Button.addEventListener('click', (e) => {
        // Prevent default behavior
        e.preventDefault();
        
        console.log('Back to Step 2 clicked');
        
        // Save current step data before moving back
        this.saveCurrentInquiryData();
        
        this.currentStep = 2;
        this.validateAttempted = false;
        this.render();
      });
    }
    
    const submitButton = this.domElement.querySelector('#submitBtn');
    if (submitButton) {
      submitButton.addEventListener('click', (e) => {
        // Prevent default behavior
        e.preventDefault();
        
        console.log('Submit button clicked');
        
        // Make sure to save current data before validation
        this.saveCurrentInquiryData();
        
        if (this.validateStep3()) {
          // Data is already saved by saveCurrentInquiryData() above
          this.submitForm();
        }
      });
    }
    
    const newInquiryButton = this.domElement.querySelector('#newInquiryBtn');
    if (newInquiryButton) {
      newInquiryButton.addEventListener('click', () => {
        console.log('New Inquiry button clicked');
        this.resetForm();
      });
    }
    
    // Setup country search and selection
    this.setupCountrySelector();
    
    // Setup regulators selection if in step 2
    if (this.currentStep === 2) {
      this.setupRegulatorSelector();
      
      // Handle primary business area change
      const primaryBusinessAreasSelect = this.domElement.querySelector('#primaryBusinessAreas') as HTMLSelectElement;
      if (primaryBusinessAreasSelect) {
        primaryBusinessAreasSelect.addEventListener('change', () => {
          console.log('Primary business area changed to:', primaryBusinessAreasSelect.value);
          this.formData.primaryBusinessAreas = primaryBusinessAreasSelect.value;
          
          // Clear the product/service category when business area changes
          this.formData.productServiceCategory = '';
          
          // Re-render to update the product service category dropdown
          this.render();
          
          // Setup event listener for product service category after render
          setTimeout(() => {
            const productServiceCategorySelectAfterRender = this.domElement.querySelector('#productServiceCategory') as HTMLSelectElement;
            if (productServiceCategorySelectAfterRender) {
              console.log('Setting up product service category change listener');
              productServiceCategorySelectAfterRender.addEventListener('change', () => {
                console.log('Product service category changed to:', productServiceCategorySelectAfterRender.value);
                this.formData.productServiceCategory = productServiceCategorySelectAfterRender.value;
                
                // Re-render if "Other" is selected to show the additional field
                if (productServiceCategorySelectAfterRender.value === 'Other') {
                  this.render();
                }
              });
            }
          }, 100);
        });
      }
      
      // Setup event listener for product service category
      const productServiceCategorySelect = this.domElement.querySelector('#productServiceCategory') as HTMLSelectElement;
      if (productServiceCategorySelect) {
        console.log('Setting up product service category change listener');
        productServiceCategorySelect.addEventListener('change', () => {
          console.log('Product service category changed to:', productServiceCategorySelect.value);
          this.formData.productServiceCategory = productServiceCategorySelect.value;
          
          // Re-render if "Other" is selected to show the additional field
          if (productServiceCategorySelect.value === 'Other') {
            this.render();
          }
        });
      }
      
      // Handle operational status change
      const operationalStatusYes = this.domElement.querySelector('#operationalStatusYes') as HTMLInputElement;
      const operationalStatusNo = this.domElement.querySelector('#operationalStatusNo') as HTMLInputElement;
      
      if (operationalStatusYes) {
        operationalStatusYes.addEventListener('change', () => {
          console.log('Operational status changed to Yes');
          if (operationalStatusYes.checked) {
            // Save current form data
            this.saveCurrentIndustryData();
            this.formData.operationalStatus = true;
          }
        });
      }
      
      if (operationalStatusNo) {
        operationalStatusNo.addEventListener('change', () => {
          console.log('Operational status changed to No');
          if (operationalStatusNo.checked) {
            // Save current form data
            this.saveCurrentIndustryData();
            this.formData.operationalStatus = false;
          }
        });
      }
      
      // Handle regulatory status change
      const regulatoryStatusYes = this.domElement.querySelector('#regulatoryStatusYes') as HTMLInputElement;
      const regulatoryStatusNo = this.domElement.querySelector('#regulatoryStatusNo') as HTMLInputElement;
      
      if (regulatoryStatusYes) {
        regulatoryStatusYes.addEventListener('change', () => {
          console.log('Regulatory status changed to Yes');
          if (regulatoryStatusYes.checked) {
            // Save current form data before changing regulatory status
            this.saveCurrentIndustryData();
            
            this.formData.regulatoryStatus = true;
            this.render();
            // Re-setup handlers after render
            this.setupRegulatorSelector();
            this.setButtonHandlers();
          }
        });
      }
      
      if (regulatoryStatusNo) {
        regulatoryStatusNo.addEventListener('change', () => {
          console.log('Regulatory status changed to No');
          if (regulatoryStatusNo.checked) {
            // Save current form data before changing regulatory status
            this.saveCurrentIndustryData();
            
            this.formData.regulatoryStatus = false;
            // Clear regulator selection when No is selected
            this.formData.regulators = [];
            this.render();
            this.setButtonHandlers();
          }
        });
      }
    }
    
    // Setup question management and file upload if in step 3
    if (this.currentStep === 3) {
      // Add question button
      const addQuestionBtn = this.domElement.querySelector('#addQuestionBtn');
      if (addQuestionBtn) {
        addQuestionBtn.addEventListener('click', () => {
          console.log('Add question button clicked');
          
          // Save current section data before re-rendering
          this.saveCurrentInquiryData();
          
          // Add new question
          this.formData.questions.push('');
          this.render();
          this.setupRemoveQuestionButtons();
        });
      }
      
      // Remove question buttons
      this.setupRemoveQuestionButtons();
      
      // File upload
      const uploadBtn = this.domElement.querySelector('#uploadBtn');
      const fileInput = this.domElement.querySelector('#fileUpload') as HTMLInputElement;
      
      if (uploadBtn && fileInput) {
        uploadBtn.addEventListener('click', (e) => {
          // Prevent default behavior
          e.preventDefault();
          
          console.log('Upload button clicked');
          
          // Save current form data before handling files
          this.saveCurrentInquiryData();
          
          if (fileInput.files && fileInput.files.length > 0) {
            const maxFiles = 5;
            const maxFileSize = 10 * 1024 * 1024; // 10MB
            const allowedTypes = ['application/pdf', 'application/msword', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 
                                 'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                 'image/png', 'image/jpeg'];
            
            // Validate number of files
            if (this.formData.files.length + fileInput.files.length > maxFiles) {
              alert(`You can upload a maximum of ${maxFiles} files.`);
              return;
            }
            
            let invalidFiles = false;
            
            // Add files to the form data
            for (let i = 0; i < fileInput.files.length; i++) {
              const file = fileInput.files[i];
              
              // Validate file size
              if (file.size > maxFileSize) {
                alert(`File "${file.name}" exceeds the maximum size of 10MB.`);
                invalidFiles = true;
                continue;
              }
              
              // Validate file type
              if (allowedTypes.indexOf(file.type) === -1) {
                alert(`File "${file.name}" is not an allowed file type.`);
                invalidFiles = true;
                continue;
              }
              
              this.formData.files.push(file);
            }
            
            if (!invalidFiles) {
              // Clear the file input properly (cross-browser solution)
              this.resetFileInput(fileInput);
            }
            
            // Re-render the file list
            this.render();
            this.setupRemoveFileButtons();
          }
        });
      }
      
      // Remove file buttons
      this.setupRemoveFileButtons();
    }
  }
  
  private setupCountrySelector(): void {
    console.log('Setting up country selector - DEBUG');
    console.log('Current countries of operation:', this.formData.countriesOfOperation);
    
    // Setup the country search
    const countrySearch = this.domElement.querySelector('#countrySearch') as HTMLInputElement;
    const countryItems = this.domElement.querySelectorAll('.country-item');
    
    if (countrySearch) {
      console.log('Found country search field');
      countrySearch.addEventListener('input', () => {
        const searchValue = countrySearch.value.toLowerCase();
        console.log('Searching for country:', searchValue);
        
        for (let i = 0; i < countryItems.length; i++) {
          const item = countryItems[i] as HTMLElement;
          const label = item.querySelector('label');
          const countryName = label ? label.textContent.toLowerCase() : '';
          
          if (countryName.indexOf(searchValue) !== -1) {
            item.style.display = '';
          } else {
            item.style.display = 'none';
          }
        }
      });
    } else {
      console.log('Country search field not found');
    }
    
    // Setup checkbox for country selection
    const countryCheckboxes = this.domElement.querySelectorAll('.country-checkbox');
    console.log('Found country checkboxes:', countryCheckboxes.length);
    
    for (let i = 0; i < countryCheckboxes.length; i++) {
      const checkbox = countryCheckboxes[i] as HTMLInputElement;
      const countryName = checkbox.value;
      
      if (this.formData.countriesOfOperation.includes(countryName)) {
        checkbox.checked = true;
      }
      
      checkbox.addEventListener('change', () => {
        console.log('Country checkbox changed:', countryName, checkbox.checked);
        
        if (checkbox.checked) {
          if (!this.formData.countriesOfOperation.includes(countryName)) {
            this.formData.countriesOfOperation.push(countryName);
            console.log('Added country:', countryName);
          }
        } else {
          const index = this.formData.countriesOfOperation.indexOf(countryName);
          if (index !== -1) {
            this.formData.countriesOfOperation.splice(index, 1);
            console.log('Removed country:', countryName);
          }
        }
        
        this.updateSelectedCountriesDisplay();
      });
    }
    
    this.setupRemoveCountryButtons();
    this.updateSelectedCountriesDisplay();
  }

  private setupRemoveCountryButtons(): void {
    console.log('Setting up remove country buttons');
    const removeButtons = this.domElement.querySelectorAll('.remove-country-btn');
    
    for (let i = 0; i < removeButtons.length; i++) {
      const button = removeButtons[i] as HTMLButtonElement;
      button.addEventListener('click', () => {
        const countryName = button.getAttribute('data-country');
        console.log('Remove country button clicked:', countryName);
        
        if (countryName) {
          const index = this.formData.countriesOfOperation.indexOf(countryName);
          if (index !== -1) {
            this.formData.countriesOfOperation.splice(index, 1);
            
            // Uncheck the corresponding checkbox
            const checkbox = this.domElement.querySelector(`.country-checkbox[value="${countryName}"]`) as HTMLInputElement;
            if (checkbox) {
              checkbox.checked = false;
            }
            
            this.updateSelectedCountriesDisplay();
          }
        }
      });
    }
  }

  private updateSelectedCountriesDisplay(): void {
    console.log('Updating selected countries display - DEBUG');
    console.log('Countries to display:', this.formData.countriesOfOperation);
    
    // Try both selectors to see which one works
    const selectedCountriesContainer = this.domElement.querySelector(`.${styles.selectedCountries}`);
    console.log('Selected countries container found with styles.selectedCountries:', !!selectedCountriesContainer);
    
    if (!selectedCountriesContainer) {
      console.warn('Could not find selected countries container with .selectedCountries class');
      console.log('Available selectors in this area:', 
        Array.from(this.domElement.querySelectorAll('.country-selector *'))
          .map(el => (el as HTMLElement).className)
          .join(', ')
      );
    }
    
    if (selectedCountriesContainer) {
      selectedCountriesContainer.innerHTML = '';
      
      if (this.formData.countriesOfOperation.length > 0) {
        console.log('Countries selected:', this.formData.countriesOfOperation.length);
        
        for (const country of this.formData.countriesOfOperation) {
          const countryElement = document.createElement('div');
          countryElement.className = 'selected-country';
          countryElement.textContent = country;
          
          const removeButton = document.createElement('button');
          removeButton.type = 'button';
          removeButton.className = 'remove-country-btn';
          removeButton.setAttribute('data-country', country);
          removeButton.innerHTML = '&times;';
          
          countryElement.appendChild(removeButton);
          selectedCountriesContainer.appendChild(countryElement);
        }
        
        // Add validation styling
        const countriesFeedback = this.domElement.querySelector('#countriesFeedback');
        if (countriesFeedback) {
          countriesFeedback.classList.remove('show-error');
        }
      } else {
        console.log('No countries selected');
        // Show validation error if validation has been attempted
        if (this.validateAttempted) {
          const countriesFeedback = this.domElement.querySelector('#countriesFeedback');
          if (countriesFeedback) {
            countriesFeedback.classList.add('show-error');
          }
        }
      }
      
      // Refresh remove buttons
      this.setupRemoveCountryButtons();
    } else {
      console.log('Selected countries container not found');
    }
  }

  private validateStep1(): boolean {
    console.log('Validating Step 1');
    this.validateAttempted = true;
    let isValid = true;
    
    const fullNameInput = this.domElement.querySelector('#fullName') as HTMLInputElement;
    const organisationNameInput = this.domElement.querySelector('#organisationName') as HTMLInputElement;
    const emailAddressInput = this.domElement.querySelector('#emailAddress') as HTMLInputElement;
    const operationLocationInput = this.domElement.querySelector('#operationLocation') as HTMLInputElement;
    const operationLengthInput = this.domElement.querySelector('#operationLength') as HTMLInputElement;
    
    // Check each field for validity
    if (!fullNameInput || !fullNameInput.value.trim()) {
      console.log('Full name is invalid');
      isValid = false;
    }
    
    if (!organisationNameInput || !organisationNameInput.value.trim()) {
      console.log('Organisation name is invalid');
      isValid = false;
    }
    
    if (!emailAddressInput || !emailAddressInput.value.trim()) {
      console.log('Email address is empty');
      isValid = false;
    } else {
      // Validate email format
      const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailPattern.test(emailAddressInput.value)) {
        console.log('Email address format is invalid');
        isValid = false;
      }
    }
    
    if (!operationLocationInput || !operationLocationInput.value.trim()) {
      console.log('Operation location is invalid');
      isValid = false;
    }
    
    if (this.formData.countriesOfOperation.length === 0) {
      console.log('Countries of operation is empty');
      isValid = false;
    }
    
    if (!operationLengthInput || !operationLengthInput.value.trim()) {
      console.log('Operation length is invalid');
      isValid = false;
    }
    
    // Re-render to show validation messages
    if (!isValid) {
      this.render();
    }
    
    console.log('Step 1 validation result:', isValid);
    return isValid;
  }

  private validateStep2(): boolean {
    console.log('Validating Step 2');
    this.validateAttempted = true;
    let isValid = true;
    
    // Primary business area validation
    const primaryBusinessArea = this.domElement.querySelector('#primaryBusinessAreas') as HTMLSelectElement;
    console.log('Primary business area:', primaryBusinessArea ? primaryBusinessArea.value : undefined);
    if (!primaryBusinessArea || !primaryBusinessArea.value) {
      console.log('Primary business area is invalid');
      isValid = false;
    } else {
      this.formData.primaryBusinessAreas = primaryBusinessArea.value;
    }
    
    // Product/service category validation
    const productServiceCategory = this.domElement.querySelector('#productServiceCategory') as HTMLSelectElement;
    console.log('Product/service category:', productServiceCategory ? productServiceCategory.value : undefined);
    if (!productServiceCategory || !productServiceCategory.value) {
      console.log('Product/service category is invalid');
      isValid = false;
    } else {
      this.formData.productServiceCategory = productServiceCategory.value;
    }
    
    // Other product/service category validation if "Other" is selected
    if (this.formData.productServiceCategory === 'Other') {
      const otherProductServiceCategory = this.domElement.querySelector('#otherProductServiceCategory') as HTMLInputElement;
      console.log('Other product/service category:', otherProductServiceCategory ? otherProductServiceCategory.value : undefined);
      if (!otherProductServiceCategory || !otherProductServiceCategory.value.trim()) {
        console.log('Other product/service category is invalid');
        isValid = false;
      } else {
        this.formData.otherProductServiceCategory = otherProductServiceCategory.value.trim();
      }
    }
    
    // Operational status validation
    const operationalStatusYes = this.domElement.querySelector('#operationalStatusYes') as HTMLInputElement;
    const operationalStatusNo = this.domElement.querySelector('#operationalStatusNo') as HTMLInputElement;
    console.log('Operational status:', 
      operationalStatusYes ? operationalStatusYes.checked : undefined, 
      operationalStatusNo ? operationalStatusNo.checked : undefined);
    if ((!operationalStatusYes || !operationalStatusYes.checked) && (!operationalStatusNo || !operationalStatusNo.checked)) {
      console.log('Operational status is invalid');
      isValid = false;
    } else {
      this.formData.operationalStatus = operationalStatusYes.checked;
    }
    
    // Regulatory status validation
    const regulatoryStatusYes = this.domElement.querySelector('#regulatoryStatusYes') as HTMLInputElement;
    const regulatoryStatusNo = this.domElement.querySelector('#regulatoryStatusNo') as HTMLInputElement;
    console.log('Regulatory status:', 
      regulatoryStatusYes ? regulatoryStatusYes.checked : undefined, 
      regulatoryStatusNo ? regulatoryStatusNo.checked : undefined);
    if ((!regulatoryStatusYes || !regulatoryStatusYes.checked) && (!regulatoryStatusNo || !regulatoryStatusNo.checked)) {
      console.log('Regulatory status is invalid');
      isValid = false;
    } else {
      this.formData.regulatoryStatus = regulatoryStatusYes.checked;
    }
    
    // Regulators validation if regulatory status is "Yes"
    if (this.formData.regulatoryStatus === true) {
      console.log('Regulators:', this.formData.regulators);
      if (this.formData.regulators.length === 0) {
        const regulatorCheckboxes = this.domElement.querySelectorAll('.regulator-checkbox:checked') as NodeListOf<HTMLInputElement>;
        console.log('Regulator checkboxes found:', regulatorCheckboxes.length);
        
        // Check if there are checked checkboxes but the array is empty (sync issue)
        if (regulatorCheckboxes.length > 0) {
          // Sync the form data
          this.formData.regulators = Array.from(regulatorCheckboxes).map(checkbox => checkbox.value);
          console.log('Synced regulators:', this.formData.regulators);
        } else {
          console.log('No regulators selected');
          isValid = false;
        }
      }
    }
    
    // Other regulator validation if "Other" is selected
    if (this.formData.regulators.indexOf('Other') !== -1) {
      const otherRegulatorInput = this.domElement.querySelector('#otherRegulator') as HTMLInputElement;
      console.log('Other regulator:', otherRegulatorInput ? otherRegulatorInput.value : undefined);
      if (otherRegulatorInput && otherRegulatorInput.value.trim()) {
        this.formData.otherRegulator = otherRegulatorInput.value.trim();
      }
    }
    
    console.log('Step 2 validation result:', isValid);
    
    // Set validation attempted flag and re-render if not valid
    if (!isValid) {
      this.render();
    }
    
    return isValid;
  }

  private validateStep3(): boolean {
    console.log('Validating Step 3');
    this.validateAttempted = true;
    let isValid = true;
    
    // Product/service description validation
    const descriptionTextarea = this.domElement.querySelector('#productServiceDescription') as HTMLTextAreaElement;
    if (!descriptionTextarea || !descriptionTextarea.value.trim()) {
      console.log('Product/service description is invalid');
      isValid = false;
    } else {
      this.formData.productServiceDescription = descriptionTextarea.value.trim();
    }
    
    // Questions validation (at least one question required)
    const questionInputs = this.domElement.querySelectorAll('.question-input') as NodeListOf<HTMLTextAreaElement>;
    let hasValidQuestion = false;
    
    // Check if at least one question is filled
    for (let i = 0; i < questionInputs.length; i++) {
      if (questionInputs[i].value.trim()) {
        hasValidQuestion = true;
        break;
      }
    }
    
    if (!hasValidQuestion) {
      console.log('No valid questions found');
      isValid = false;
    }
    
    // FAQ confirmation validation
    const faqYes = this.domElement.querySelector('#faqConfirmationYes') as HTMLInputElement;
    const faqNo = this.domElement.querySelector('#faqConfirmationNo') as HTMLInputElement;
    if ((!faqYes || !faqYes.checked) && (!faqNo || !faqNo.checked)) {
      console.log('FAQ confirmation is invalid');
      isValid = false;
    }
    
    // Consent validation
    const consentCheckbox = this.domElement.querySelector('#consentCheckbox') as HTMLInputElement;
    if (!consentCheckbox || !consentCheckbox.checked) {
      console.log('Consent is not checked');
      isValid = false;
    }
    
    // Re-render to show validation messages if invalid
    if (!isValid) {
      console.log('Step 3 validation failed');
      this.render();
      
      // Re-setup event handlers after render
      this.setupRemoveQuestionButtons();
      this.setupRemoveFileButtons();
    }
    
    console.log('Step 3 validation result:', isValid);
    return isValid;
  }

  private isValidEmail(email: string): boolean {
    const regex = /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return regex.test(String(email).toLowerCase());
  }

  private saveStep1Data(): void {
    console.log('Saving Step 1 data');
    const fullNameInput = this.domElement.querySelector('#fullName') as HTMLInputElement;
    const organisationNameInput = this.domElement.querySelector('#organisationName') as HTMLInputElement;
    const contactNumberInput = this.domElement.querySelector('#contactNumber') as HTMLInputElement;
    const emailAddressInput = this.domElement.querySelector('#emailAddress') as HTMLInputElement;
    const websiteAddressInput = this.domElement.querySelector('#websiteAddress') as HTMLInputElement;
    const operationLocationInput = this.domElement.querySelector('#operationLocation') as HTMLInputElement;
    const operationLengthInput = this.domElement.querySelector('#operationLength') as HTMLInputElement;
    
    if (fullNameInput) this.formData.fullName = fullNameInput.value.trim();
    if (organisationNameInput) this.formData.organisationName = organisationNameInput.value.trim();
    if (contactNumberInput) this.formData.contactNumber = contactNumberInput.value.trim();
    if (emailAddressInput) this.formData.emailAddress = emailAddressInput.value.trim();
    if (websiteAddressInput) this.formData.websiteAddress = websiteAddressInput.value.trim();
    if (operationLocationInput) this.formData.operationLocation = operationLocationInput.value.trim();
    if (operationLengthInput) this.formData.operationLength = operationLengthInput.value.trim();
    
    // Note: Countries of operation are managed by event listeners which update the countriesOfOperation array
    console.log('Saved form data:', this.formData);
    
    // Move to the next step
    this.currentStep = 2;
    this.validateAttempted = false;
    this.render();
  }

  private saveStep2Data(): void {
    console.log('Saving Step 2 data');
    
    // Get primary business area
    const primaryBusinessAreasSelect = this.domElement.querySelector('#primaryBusinessAreas') as HTMLSelectElement;
    if (primaryBusinessAreasSelect) {
      console.log('Saving primary business area:', primaryBusinessAreasSelect.value);
      this.formData.primaryBusinessAreas = primaryBusinessAreasSelect.value;
    }
    
    // Get product/service category
    const productServiceCategorySelect = this.domElement.querySelector('#productServiceCategory') as HTMLSelectElement;
    if (productServiceCategorySelect) {
      console.log('Saving product/service category:', productServiceCategorySelect.value);
      this.formData.productServiceCategory = productServiceCategorySelect.value;
    }
    
    // Get other product/service category if applicable
    if (this.formData.productServiceCategory === 'Other') {
      const otherProductServiceCategoryInput = this.domElement.querySelector('#otherProductServiceCategory') as HTMLInputElement;
      if (otherProductServiceCategoryInput) {
        console.log('Saving other product/service category:', otherProductServiceCategoryInput.value);
        this.formData.otherProductServiceCategory = otherProductServiceCategoryInput.value;
      }
    }
    
    // Get operational status
    const operationalStatusYes = this.domElement.querySelector('#operationalStatusYes') as HTMLInputElement;
    const operationalStatusNo = this.domElement.querySelector('#operationalStatusNo') as HTMLInputElement;
    if (operationalStatusYes && operationalStatusNo) {
      console.log('Saving operational status:', operationalStatusYes.checked ? 'Yes' : 'No');
      if (operationalStatusYes.checked) {
        this.formData.operationalStatus = true;
      } else if (operationalStatusNo.checked) {
        this.formData.operationalStatus = false;
      }
    }
    
    // Get regulatory status
    const regulatoryStatusYes = this.domElement.querySelector('#regulatoryStatusYes') as HTMLInputElement;
    const regulatoryStatusNo = this.domElement.querySelector('#regulatoryStatusNo') as HTMLInputElement;
    if (regulatoryStatusYes && regulatoryStatusNo) {
      console.log('Saving regulatory status:', regulatoryStatusYes.checked ? 'Yes' : 'No');
      if (regulatoryStatusYes.checked) {
        this.formData.regulatoryStatus = true;
        
        // Check for regulator checkboxes
        const regulatorCheckboxes = this.domElement.querySelectorAll('.regulator-checkbox:checked') as NodeListOf<HTMLInputElement>;
        if (regulatorCheckboxes && regulatorCheckboxes.length > 0) {
          // Update the regulators array directly from checked checkboxes
          this.formData.regulators = Array.from(regulatorCheckboxes).map(checkbox => checkbox.value);
          console.log('Saving regulators:', this.formData.regulators);
        }
      } else if (regulatoryStatusNo.checked) {
        this.formData.regulatoryStatus = false;
      }
    }
    
    // Get other regulator if applicable
    if (this.formData.regulators.indexOf('Other') !== -1) {
      const otherRegulatorInput = this.domElement.querySelector('#otherRegulator') as HTMLInputElement;
      console.log('Saving other regulator:', otherRegulatorInput ? otherRegulatorInput.value : undefined);
      if (otherRegulatorInput && otherRegulatorInput.value.trim()) {
        this.formData.otherRegulator = otherRegulatorInput.value.trim();
      }
    }
    
    console.log('Step 2 data saved:', this.formData);
  }

  private saveStep3Data(): void {
    // Get product/service description
    const descriptionTextarea = this.domElement.querySelector('#productServiceDescription') as HTMLTextAreaElement;
    if (descriptionTextarea) {
      this.formData.productServiceDescription = descriptionTextarea.value;
    }
    
    // Get questions
    const questionInputs = this.domElement.querySelectorAll('.question-input') as NodeListOf<HTMLTextAreaElement>;
    this.formData.questions = [];
    
    for (let i = 0; i < questionInputs.length; i++) {
      const question = questionInputs[i].value.trim();
      if (question) {
        this.formData.questions.push(question);
      }
    }
    
    // Get additional information
    const additionalInfoTextarea = this.domElement.querySelector('#additionalInformation') as HTMLTextAreaElement;
    if (additionalInfoTextarea) {
      this.formData.additionalInformation = additionalInfoTextarea.value;
    }
    
    // Get FAQ confirmation
    const faqYes = this.domElement.querySelector('#faqConfirmationYes') as HTMLInputElement;
    const faqNo = this.domElement.querySelector('#faqConfirmationNo') as HTMLInputElement;
    if (faqYes && faqNo) {
      if (faqYes.checked) {
        this.formData.faqConfirmation = true;
      } else if (faqNo.checked) {
        this.formData.faqConfirmation = false;
      }
    }
    
    // Get consent confirmation
    const consentCheckbox = this.domElement.querySelector('#consentCheckbox') as HTMLInputElement;
    if (consentCheckbox) {
      this.formData.consentConfirmation = consentCheckbox.checked;
    }
    
    // Files are handled by the file upload event handler
  }

  private submitForm(): void {
    // Show loading state
    console.log('Starting form submission process...');
    const submitButton = this.domElement.querySelector('.submit-btn') as HTMLButtonElement;
    if (submitButton) {
      submitButton.disabled = true;
      submitButton.innerHTML = 'Submitting...';
    }
    
    console.log('Form data being submitted:', JSON.stringify(this.formData, null, 2));
    
    // First, create the basic information list item
    console.log('Attempting to submit Basic Information list item...');
    this.createBasicInformationListItem()
      .then((basicInfoResponse) => {
        console.log('Basic Information API response status:', basicInfoResponse.status, basicInfoResponse.statusText);
        // Get the ID of the created item
        return basicInfoResponse.json().then(data => {
          console.log('Basic Information submitted successfully, response data:', data);
          
          // Use the ID from the first submission for the related lists
          const basicInfoId = data.Id || data.id;
          console.log('Using Basic Information ID for related lists:', basicInfoId);
          
          // Submit to Industry Information list
          console.log('Attempting to submit Industry Information list item...');
          return this.createIndustryInformationListItem(basicInfoId)
            .then((industryInfoResponse) => {
              console.log('Industry Information API response status:', industryInfoResponse.status, industryInfoResponse.statusText);
              console.log('Industry Information submitted successfully');
              
              // Submit to Enquiry Details list
              console.log('Attempting to submit Enquiry Details list item...');
              return this.createEnquiryDetailsListItem(basicInfoId);
            });
        });
      })
      .then(() => {
        console.log('All form data submitted successfully to SharePoint lists');
        
        // Upload files if there are any
        if (this.formData.files && this.formData.files.length > 0) {
          console.log('Starting file upload process for', this.formData.files.length, 'files');
          return this.uploadFiles(this.formData.files);
        }
        console.log('No files to upload');
        return Promise.resolve();
      })
      .then(() => {
        console.log('Form submission completed successfully, showing thank you step');
        // Show thank you step
        this.currentStep = 4;
        this.render();
      })
      .catch((error) => {
        console.error('Error submitting form:', error);
        if (error.response) {
          console.error('Response status:', error.response.status);
          error.response.json().then(errorData => {
            console.error('Detailed error information:', errorData);
          }).catch(e => {
            console.error('Could not parse error response as JSON');
          });
        }
        if (error.message) {
          console.error('Error message:', error.message);
        }
        if (error.stack) {
          console.error('Error stack:', error.stack);
        }
        
        // Re-enable submit button
        if (submitButton) {
          submitButton.disabled = false;
          submitButton.innerHTML = this.properties.submitButtonText || 'Submit';
        }
        
        // Show error message to the user
        const errorContainer = this.domElement.querySelector('.error-message');
        if (errorContainer) {
          errorContainer.innerHTML = 'There was an error submitting your form. Please try again later.';
          errorContainer.className = 'error-message visible';
        }
      });
  }

  /**
   * Creates a list item in the Basic Information list
   */
  private createBasicInformationListItem(): Promise<SPHttpClientResponse> {
    const url = 'https://www.ifwg.co.za/_api/web/lists/getbytitle(\'Basic Information\')/items';
    console.log('Basic Information API URL:', url);
    
    const listItem = {
      Title: this.formData.fullName, // Title is a required field in SharePoint lists
      FullName: this.formData.fullName,
      OrganisationName: this.formData.organisationName,
      ContactNumber: this.formData.contactNumber,
      EmailAddress: this.formData.emailAddress,
      WebsiteAddress: this.formData.websiteAddress,
      OperationLocation: this.formData.operationLocation,
      CountriesOfOperation: this.formData.countriesOfOperation.join(', '),
      OperationLength: this.formData.operationLength,
      SubmissionDate: new Date().toISOString()
    };
    console.log('Basic Information list item data:', JSON.stringify(listItem, null, 2));
    
    const headers = {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=nometadata',
      'odata-version': ''
    };
    console.log('Request headers:', headers);
    
    return this.context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: headers,
        body: JSON.stringify(listItem)
      }
    );
  }

  /**
   * Creates a list item in the Industry Information list
   */
  private createIndustryInformationListItem(basicInfoId: number): Promise<SPHttpClientResponse> {
    const url = 'https://www.ifwg.co.za/_api/web/lists/getbytitle(\'Industry Information\')/items';
    console.log('Industry Information API URL:', url);
    
    const listItem = {
      Title: `Industry Info for ${this.formData.fullName}`, // Title is a required field in SharePoint lists
      BasicInformationId: basicInfoId.toString(), // Reference to the Basic Information list item
      PrimaryBusinessAreas: this.formData.primaryBusinessAreas,
      ProductServiceCategory: this.formData.productServiceCategory,
      OtherProductServiceCategory: this.formData.otherProductServiceCategory,
      OperationalStatus: this.formData.operationalStatus === true ? 'Yes' : (this.formData.operationalStatus === false ? 'No' : ''),
      RegulatoryStatus: this.formData.regulatoryStatus === true ? 'Yes' : (this.formData.regulatoryStatus === false ? 'No' : ''),
      Regulators: this.formData.regulators.join(', '),
      OtherRegulator: this.formData.otherRegulator
    };
    console.log('Industry Information list item data:', JSON.stringify(listItem, null, 2));
    
    return this.context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: JSON.stringify(listItem)
      }
    );
  }

  /**
   * Creates a list item in the Enquiry Details list
   */
  private createEnquiryDetailsListItem(basicInfoId: number): Promise<SPHttpClientResponse> {
    const url = 'https://www.ifwg.co.za/_api/web/lists/getbytitle(\'Enquiry Details\')/items';
    console.log('Enquiry Details API URL:', url);
    
    const listItem = {
      Title: `Enquiry for ${this.formData.fullName}`, // Title is a required field in SharePoint lists
      BasicInformationId: basicInfoId.toString(), // Reference to the Basic Information list item
      ProductServiceDescription: this.formData.productServiceDescription,
      Questions: this.formData.questions.filter(q => q.trim() !== '').join('\n\n'),
      AdditionalInformation: this.formData.additionalInformation,
      FAQConfirmation: this.formData.faqConfirmation === true ? 'Yes' : (this.formData.faqConfirmation === false ? 'No' : ''),
      ConsentConfirmation: this.formData.consentConfirmation ? 'Yes' : 'No',
      FileAttachments: this.formData.files.length > 0 ? 'Yes' : 'No',
      NumberOfAttachments: this.formData.files.length.toString()
    };
    console.log('Enquiry Details list item data:', JSON.stringify(listItem, null, 2));
    
    return this.context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: JSON.stringify(listItem)
      }
    );
  }

  /**
   * Uploads files to the document library
   */
  private uploadFiles(files: File[]): Promise<any> {
    // If no files, return resolved promise
    if (!files || files.length === 0) {
      return Promise.resolve();
    }
    
    // Create folder with timestamp to group files
    const timestamp = new Date().getTime();
    const folderName = `Enquiry_${this.formData.fullName.replace(/\s+/g, '_')}_${timestamp}`;
    console.log('Creating folder for file uploads:', folderName);
    
    // Upload each file and return promise that resolves when all uploads are complete
    const uploadPromises = files.map((file, index) => {
      console.log(`Starting upload for file ${index + 1}/${files.length}: ${file.name}`);
      return this.uploadFile(file, folderName);
    });
    return Promise.all(uploadPromises);
  }

  /**
   * Uploads a single file to the document library
   */
  private uploadFile(file: File, folderName: string): Promise<SPHttpClientResponse> {
    // Create the folder first
    console.log(`Creating folder ${folderName} for file ${file.name}`);
    return this.createFolder(folderName)
      .then(() => {
        // Now upload the file to the folder
        const url = `https://www.ifwg.co.za/_api/web/getfolderbyserverrelativeurl('/EnquiryFormDocuments/${folderName}')/files/add(url='${file.name}',overwrite=true)`;
        console.log('File upload URL:', url);
        
        return this.context.spHttpClient.post(
          url,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/octet-stream',
              'odata-version': ''
            },
            body: file
          }
        );
      })
      .then(response => {
        console.log(`File ${file.name} upload response status:`, response.status, response.statusText);
        return response;
      })
      .catch(error => {
        console.error(`Error uploading file ${file.name}:`, error);
        throw error;
      });
  }

  /**
   * Creates a folder in the document library
   */
  private createFolder(folderName: string): Promise<SPHttpClientResponse> {
    const url = `https://www.ifwg.co.za/_api/web/folders`;
    console.log('Folder creation URL:', url);
    
    const folderUrl = `/EnquiryFormDocuments/${folderName}`;
    console.log('Creating folder with ServerRelativeUrl:', folderUrl);
    
    return this.context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: JSON.stringify({
          'ServerRelativeUrl': folderUrl
        })
      }
    )
      .then(response => {
        console.log('Folder creation response status:', response.status, response.statusText);
        return response;
      })
      .catch(error => {
        console.error('Error creating folder:', error);
        throw error;
      });
  }

  private resetForm(): void {
    this.currentStep = 1;
    this.formData = {
      // Basic Information
      fullName: '',
      organisationName: '',
      contactNumber: '',
      emailAddress: '',
      websiteAddress: '',
      operationLocation: '',
      countriesOfOperation: [],
      operationLength: '',
      
      // Industry Information
      primaryBusinessAreas: '',
      productServiceCategory: '',
      otherProductServiceCategory: '',
      operationalStatus: null,
      regulatoryStatus: null,
      regulators: [],
      otherRegulator: '',
      
      // Enquiry Details
      productServiceDescription: '',
      questions: [''],
      additionalInformation: '',
      faqConfirmation: null,
      consentConfirmation: false,
      
      // Attachments
      files: []
    };
    this.render();
  }

  private setupRegulatorSelector(): void {
    const regulatorCheckboxes = this.domElement.querySelectorAll('.regulator-checkbox') as NodeListOf<HTMLInputElement>;
    
    for (let i = 0; i < regulatorCheckboxes.length; i++) {
      const checkbox = regulatorCheckboxes[i];
      
      checkbox.addEventListener('change', () => {
        const regulator = checkbox.value;
        const index = this.formData.regulators.indexOf(regulator);
        
        if (checkbox.checked && index === -1) {
          // Add regulator to the list
          this.formData.regulators.push(regulator);
        } else if (!checkbox.checked && index !== -1) {
          // Remove regulator from the list
          this.formData.regulators.splice(index, 1);
        }
        
        // If Other is checked/unchecked, re-render to show/hide the text field
        if (regulator === 'Other') {
          this.render();
          // Re-setup regulator selector after render if Other is checked
          if (checkbox.checked) {
            this.setupRegulatorSelector();
          }
        }
      });
    }
  }

  private setupRemoveQuestionButtons(): void {
    const removeButtons = this.domElement.querySelectorAll('.remove-question-btn');
    
    for (let i = 0; i < removeButtons.length; i++) {
      const button = removeButtons[i] as HTMLButtonElement;
      
      button.addEventListener('click', () => {
        const index = parseInt(button.getAttribute('data-index'), 10);
        if (!isNaN(index) && index > 0 && index < this.formData.questions.length) {
          // Save current form data before modifying and re-rendering
          this.saveCurrentInquiryData();
          
          // Remove the question at the specified index
          this.formData.questions.splice(index, 1);
          this.render();
          this.setupRemoveQuestionButtons();
        }
      });
    }
  }
  
  private setupRemoveFileButtons(): void {
    const removeButtons = this.domElement.querySelectorAll('.remove-file-btn');
    
    for (let i = 0; i < removeButtons.length; i++) {
      const button = removeButtons[i] as HTMLButtonElement;
      
      button.addEventListener('click', (e) => {
        // Prevent default behavior
        e.preventDefault();
        
        // Save current form data
        this.saveCurrentInquiryData();
        
        const index = parseInt(button.getAttribute('data-index'), 10);
        if (!isNaN(index) && index >= 0 && index < this.formData.files.length) {
          this.formData.files.splice(index, 1);
          this.render();
          this.setupRemoveFileButtons();
        }
      });
    }
  }

  private saveCurrentInquiryData(): void {
    console.log('Saving current inquiry data without changing steps');
    
    // Get product/service description
    const descriptionTextarea = this.domElement.querySelector('#productServiceDescription') as HTMLTextAreaElement;
    if (descriptionTextarea) {
      this.formData.productServiceDescription = descriptionTextarea.value;
    }
    
    // Get existing questions (preserve user input)
    const questionInputs = this.domElement.querySelectorAll('.question-input') as NodeListOf<HTMLTextAreaElement>;
    
    // Only update existing questions, don't change the array length
    for (let i = 0; i < questionInputs.length; i++) {
      if (i < this.formData.questions.length) {
        this.formData.questions[i] = questionInputs[i].value;
      }
    }
    
    // Get additional information
    const additionalInfoTextarea = this.domElement.querySelector('#additionalInformation') as HTMLTextAreaElement;
    if (additionalInfoTextarea) {
      this.formData.additionalInformation = additionalInfoTextarea.value;
    }
    
    // Get FAQ confirmation
    const faqYes = this.domElement.querySelector('#faqConfirmationYes') as HTMLInputElement;
    const faqNo = this.domElement.querySelector('#faqConfirmationNo') as HTMLInputElement;
    if (faqYes && faqNo) {
      if (faqYes.checked) {
        this.formData.faqConfirmation = true;
      } else if (faqNo.checked) {
        this.formData.faqConfirmation = false;
      }
    }
    
    // Get consent confirmation
    const consentCheckbox = this.domElement.querySelector('#consentCheckbox') as HTMLInputElement;
    if (consentCheckbox) {
      this.formData.consentConfirmation = consentCheckbox.checked;
    }
    
    console.log('Current inquiry data saved:', this.formData);
  }

  /**
   * Safely resets a file input element across different browsers
   */
  private resetFileInput(fileInput: HTMLInputElement): void {
    try {
      // Clear the value (works in most browsers)
      fileInput.value = '';
      
      // For IE/Edge, try creating a new form and reset
      if (!fileInput.value) {
        return; // If successful, we're done
      }
      
      // If the above method failed, try cloning and replacing
      const parentNode = fileInput.parentNode;
      if (parentNode) {
        const newInput = fileInput.cloneNode(true) as HTMLInputElement;
        newInput.value = '';
        parentNode.replaceChild(newInput, fileInput);
        
        // Reattach event listeners if needed
        this.setButtonHandlers();
      }
    } catch (error) {
      console.error('Error resetting file input:', error);
    }
  }

  /**
   * Saves the current industry form data without moving to another step
   */
  private saveCurrentIndustryData(): void {
    console.log('Saving current industry step data without changing steps');
    
    // Get primary business area
    const primaryBusinessArea = this.domElement.querySelector('#primaryBusinessAreas') as HTMLSelectElement;
    if (primaryBusinessArea) {
      this.formData.primaryBusinessAreas = primaryBusinessArea.value;
    }
    
    // Get product/service category
    const productServiceCategory = this.domElement.querySelector('#productServiceCategory') as HTMLSelectElement;
    if (productServiceCategory) {
      this.formData.productServiceCategory = productServiceCategory.value;
    }
    
    // Get other product/service category if applicable
    if (this.formData.productServiceCategory === 'Other') {
      const otherProductServiceCategory = this.domElement.querySelector('#otherProductServiceCategory') as HTMLInputElement;
      if (otherProductServiceCategory) {
        this.formData.otherProductServiceCategory = otherProductServiceCategory.value;
      }
    }
    
    // Get operational status - IMPORTANT: Preserve this across re-renders
    const operationalStatusYes = this.domElement.querySelector('#operationalStatusYes') as HTMLInputElement;
    const operationalStatusNo = this.domElement.querySelector('#operationalStatusNo') as HTMLInputElement;
    if (operationalStatusYes && operationalStatusNo) {
      if (operationalStatusYes.checked) {
        this.formData.operationalStatus = true;
      } else if (operationalStatusNo.checked) {
        this.formData.operationalStatus = false;
      }
    }
    
    // Regulatory status is handled by the callers
    
    // Get regulators if applicable
    if (this.formData.regulatoryStatus === true) {
      const regulatorCheckboxes = this.domElement.querySelectorAll('.regulator-checkbox:checked') as NodeListOf<HTMLInputElement>;
      if (regulatorCheckboxes && regulatorCheckboxes.length > 0) {
        this.formData.regulators = Array.from(regulatorCheckboxes).map(checkbox => checkbox.value);
      }
    }
    
    // Get other regulator if applicable
    if (this.formData.regulators.indexOf('Other') !== -1) {
      const otherRegulator = this.domElement.querySelector('#otherRegulator') as HTMLInputElement;
      if (otherRegulator) {
        this.formData.otherRegulator = otherRegulator.value;
      }
    }
    
    console.log('Current industry data saved:', this.formData);
  }

  protected getDataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('formTitle', {
                  label: 'Form Title'
                }),
                PropertyPaneTextField('submitButtonText', {
                  label: 'Submit Button Text'
                }),
                PropertyPaneTextField('thankYouMessage', {
                  label: 'Thank You Message'
                }),
                PropertyPaneTextField('faqPageUrl', {
                  label: 'FAQs Page URL'
                }),
                PropertyPaneTextField('submissionListName', {
                  label: 'Submission List Name'
                }),
                PropertyPaneTextField('documentLibraryName', {
                  label: 'Document Library Name'
                }),
                PropertyPaneTextField('notificationEmail', {
                  label: 'Notification Email'
                }),
                PropertyPaneTextField('adminGroupName', {
                  label: 'Admin Group Name'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
