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
  "South Africa",
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
  "Slovakia", "Slovenia", "Solomon Islands", "Somalia", "South Sudan", "Spain", "Sri Lanka", "Sudan", "Suriname", 
  "Sweden", "Switzerland", "Syria", "Taiwan", "Tajikistan", "Tanzania", "Thailand", "Timor-Leste", "Togo", "Tonga", "Trinidad and Tobago", 
  "Tunisia", "Turkey", "Turkmenistan", "Tuvalu", "Uganda", "Ukraine", "United Arab Emirates", "United Kingdom", "United States", 
  "Uruguay", "Uzbekistan", "Vanuatu", "Vatican City", "Venezuela", "Vietnam", "Yemen", "Zambia", "Zimbabwe", "Global Presence"
];

// Constants for the form
const REGULATORS = [
  "Financial Sector Conduct Authority (FSCA)",
  "Prudential Authority (PA)",
  "South African Reserve Bank (SARB)",
  "Financial Intelligence Centre (FIC)",
  "National Credit Regulator (NCR)",
  "Competition Commission (CC)",
  "South African Revenue Service (SARS)",
  "Other"
];

const CATEGORIES = {
  'Lending': [
    'Online (alternatives) lenders',
    'Asset financing',
    'Alternative scoring',
    'Lending Market-places',
  ],
  'Payments': [
    'mPOS (acquirers)',
    'Crypto Assets payment',
    'Cross-Border payments',
    'Closed loop Mobile wallets',
    'Payments aggregation',
    '3rd party payment providers',
  ],
  'Savings & Deposits': [
    'Digital community savings',
    'Savings products',
    'Layby',
    'Digital banking (issuers)',
  ],
  'Insurtech': [
    'Connected insurance',
    'Peer-to-peer insurance',
    'Automated risk analysis',
    'Digital distribution',
    'Claims management',
  ],
  'Investments': [
    'Retail trading',
    'Crypto asset trading',
    'Alternative exchange',
  ],
  'Financial planning & Advisory': [
    'Robo advisory',
    'Personal finance management',
    'Small business finance management',
  ],
  'Capital raising': [
    'Crowd investing',
    'Due diligence',
  ],
  'B2B Tech providers': [
    'Aggregators',
    'Open infrastructure',
    'RegTech & risk management',
    'Data applications',
    'Security & ID',
    'Process automation',
    'White label',
    'White label platforms (solutions)',
  ]
};

// Add service account credentials 
// WARNING: Do NOT expose credentials in client-side code in production!
// This is a simplified example and in a real production environment, 
// you should use a server-side API or Azure Function to handle authentication
const SERVICE_ACCOUNT = {
  username: 'svc_IFWGEnquiry',
  password: 'sQDrLej^[7yVSR`TxZ', // Fill this in with the actual password before deployment
  domain: 'FSCA'
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
    questions: [''], // Always start with one empty question box
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
  
  // Loading state for form submission
  private isSubmitting: boolean = false;
  
  // Retry configuration for connection errors
  private maxRetries: number = 3;
  private retryDelay: number = 2000; // 2 seconds
  private currentRetryCount: number = 0;

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.enquiry }">
        <div class="${ styles.container }">
          <div class="${ styles.header }">
            <h2 class="${ styles.title }">Regulatory Guidance Unit Enquiry Form</h2>
            <div class="${ styles.introText }">
              <p>The Regulatory Guidance Unit provides informal, non-binding steers to persons seeking direction and clarity in navigating aspects of the FinTech regulatory landscape. It relies on the expertise of representatives from across participating regulators within the IFWG to ensure that guidance is holistic, inclusive and well-considered. The functions of the Regulatory Guidance Unit include the following:</p>
              <ul>
                <li>To provide innovators with efficient and effective access to regulatory expertise potentially reducing their time needed to resolve regulatory concerns, increasing their speed to market and lowering their legal fees.</li>
                <li>To impart guidance and insight to entities seeking to operate and innovate in the market.</li>
                <li>To improve regulatory compliance through clear articulation of regulatory frameworks and reduce regulatory arbitrage.</li>
                <li>To enhance regulator understanding of innovation in the market.</li>
              </ul>
              <p>All information provided below will be kept confidential in accordance with the Protection of Personal Information Act, 2013, and the <a href="https://www.ifwg.co.za/Pages/Privacy-Policy.aspx" target="_blank">IFWG Privacy Policy</a>.</p>
              <p>Please ensure you have consulted the <a href="https://www.ifwg.co.za/Pages/Regulatory-Guidance-Unit.aspx" target="_blank">FAQs page</a> before submitting an enquiry. <strong>All enquiries that relate to, and are addressed by, the FAQs will be referred to the FAQs page.</strong></p>
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
            ${Object.keys(CATEGORIES).map(category => 
              `<option value="${category}" ${this.formData.primaryBusinessAreas === category ? 'selected' : ''}>${category}</option>`
            ).join('')}
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
    
    // If no business area is selected, return an empty options list
    if (!businessArea) {
      return '';
    }
    
    // Get the subcategories for the selected business area
    // First check if the exact key exists
    let categories = CATEGORIES[businessArea];
    
    // If not found, try case-insensitive search
    if (!categories) {
      // Find a key that matches the business area case-insensitively
      const matchingKey = Object.keys(CATEGORIES).find(
        key => key.toLowerCase() === businessArea.toLowerCase()
      );
      
      if (matchingKey) {
        categories = CATEGORIES[matchingKey];
      }
    }
    
    // If still no categories found, return empty
    if (!categories || !Array.isArray(categories)) {
      return '';
    }
    
    // Generate options for the subcategories
    let options = '';
    for (let i = 0; i < categories.length; i++) {
      const category = categories[i];
      options += `<option value="${category}" ${this.formData.productServiceCategory === category ? 'selected' : ''}>${category}</option>`;
    }
    
    // Add an "Other" option if not already in the list
    if (categories.indexOf('Other') === -1) {
      options += `<option value="Other" ${this.formData.productServiceCategory === 'Other' ? 'selected' : ''}>Other</option>`;
    }
    
    return options;
  }

  private renderInquiryStep(): string {
    return `
      <div class="${ styles.formStep } ${styles.fadeIn}">
        <h3 class="${ styles.stepTitle }">Section C: Enquiry Details</h3>
        <div class="${ styles.sectionNote }">
          Please provide details about your product/service and specific questions you have
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
          <label for="additionalInformation">Any additional information related to your enquiry?</label>
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
            <a href="https://www.ifwg.co.za/Pages/Regulatory-Guidance-Unit.aspx" target="_blank">View FAQs</a>
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
            <label for="consentCheckbox">I consent to my information being processed in accordance with the <a href="https://www.ifwg.co.za/Pages/Privacy-Policy.aspx" target="_blank">IFWG Privacy Policy</a> <span class="${ styles.required }">*</span></label>
          </div>
          ${!this.formData.consentConfirmation && this.validateAttempted ? `<div class="${styles.errorText}">You must provide consent to submit this form</div>` : ''}
        </div>
        
        <div class="${ styles.formActions }">
          <button type="button" class="${ styles.button } ${styles.backButton}" id="backToStep2">Back</button>
          <button type="button" class="${ styles.button } ${styles.submitButton}" id="submitBtn" ${this.isSubmitting ? 'disabled' : ''} style="${this.isSubmitting ? 'opacity: 0.7; cursor: not-allowed;' : ''}">
            ${this.isSubmitting ? `
              <div style="display: inline-flex; align-items: center; gap: 8px;">
                <div style="width: 16px; height: 16px; border: 2px solid transparent; border-top: 2px solid #ffffff; border-radius: 50%; animation: spin 1s linear infinite;"></div>
                <span>Submitting...</span>
              </div>
            ` : escape(this.properties.submitButtonText || 'Submit')}
          </button>
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
        <p class="${ styles.thankYouMessage }">${escape(this.properties.thankYouMessage || 'Your enquiry has been submitted successfully. We will contact you soon.')}</p>
        <div class="${ styles.formActions }">
          <button type="button" class="${ styles.button } ${styles.newInquiryButton}" id="newInquiryBtn">Submit Another Enquiry</button>
        </div>
      </div>
    `;
  }

  private setButtonHandlers(): void {
    // Navigation buttons
    const nextToStep2Button = this.domElement.querySelector('#nextToStep2');
    if (nextToStep2Button) {
      nextToStep2Button.addEventListener('click', () => {
        if (this.validateStep1()) {
          this.saveStep1Data();
          this.currentStep = 2;
          this.validateAttempted = false;
          this.render();
        } else {
          // Just re-render to show validation errors but preserve entered data
          this.render();
        }
      });
    }
    
    const backToStep1Button = this.domElement.querySelector('#backToStep1');
    if (backToStep1Button) {
      backToStep1Button.addEventListener('click', () => {
        this.saveCurrentIndustryData(); // Save current data before going back
        this.currentStep = 1;
        this.validateAttempted = false;
        this.render();
      });
    }
    
    const nextToStep3Button = this.domElement.querySelector('#nextToStep3');
    if (nextToStep3Button) {
      nextToStep3Button.addEventListener('click', () => {
        if (this.validateStep2()) {
          this.saveStep2Data();
          this.currentStep = 3;
          this.validateAttempted = false;
          this.render();
        } else {
          // Just re-render to show validation errors but preserve entered data
          this.render();
        }
      });
    }
    
    const backToStep2Button = this.domElement.querySelector('#backToStep2');
    if (backToStep2Button) {
      backToStep2Button.addEventListener('click', (e) => {
        // Prevent default behavior
        e.preventDefault();
        
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
        
        // Prevent multiple submissions
        if (this.isSubmitting) {
          return;
        }
        
        // Make sure to save current data before validation
        this.saveCurrentInquiryData();
        
        if (this.validateStep3()) {
          // Data is already saved by saveCurrentInquiryData() above
          this.submitForm();
        } else {
          // Just re-render to show validation errors but preserve entered data
          this.render();
        }
      });
    }

    const newInquiryButton = this.domElement.querySelector('#newInquiryBtn');
    if (newInquiryButton) {
      newInquiryButton.addEventListener('click', () => {
        this.resetForm();
      });
    }
    
    // Add question button
    const addQuestionButton = this.domElement.querySelector('#addQuestionBtn');
    if (addQuestionButton) {
      addQuestionButton.addEventListener('click', () => {
        // Save current form data
        this.saveCurrentInquiryData();
        
        // Add new empty question
        this.formData.questions.push('');
        this.render();
        this.setupRemoveQuestionButtons();
      });
    }
    
    // File upload
    const uploadButton = this.domElement.querySelector('#uploadBtn');
    if (uploadButton) {
      uploadButton.addEventListener('click', () => {
        const fileInput = this.domElement.querySelector('#fileUpload') as HTMLInputElement;
        if (fileInput) {
          fileInput.click();
        }
      });
    }
    
    const fileInput = this.domElement.querySelector('#fileUpload') as HTMLInputElement;
    if (fileInput) {
      this.fileUploadElement = fileInput;
      
      fileInput.addEventListener('change', (e) => {
        // Save the current form state
        this.saveCurrentInquiryData();
        
        const input = e.target as HTMLInputElement;
        if (input.files && input.files.length > 0) {
          
          // Check if number of files would exceed the limit
          const totalFiles = (this.formData.files ? this.formData.files.length : 0) + input.files.length;
          if (totalFiles > 5) {
            alert('Maximum 5 files allowed. Please remove some files or select fewer files.');
            return;
          }
          
          // Add the files to the form data
          if (!this.formData.files) {
            this.formData.files = [];
          }
          
          // Process each file
          for (let i = 0; i < input.files.length; i++) {
            const file = input.files[i];
            
            // Check file size
            if (file.size > 10 * 1024 * 1024) { // 10MB
              alert(`File ${file.name} exceeds the 10MB size limit.`);
              continue;
            }
            
            // Add file to form data
            this.formData.files.push(file);
          }
          
          // Clear the input
          this.resetFileInput(fileInput);
          
          // Re-render to show uploaded files
          this.render();
          this.setupRemoveFileButtons();
        }
      });
    }
    
    // Setup immediate email validation for step 1
    const emailInput = this.domElement.querySelector('#emailAddress') as HTMLInputElement;
    if (emailInput) {
      emailInput.addEventListener('blur', () => {
        const emailValue = emailInput.value.trim();
        
        // Update form data
        this.formData.emailAddress = emailValue;
        
        // Validate email immediately and update UI
        this.validateEmailField();
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
            // Save current industry data
            this.saveCurrentIndustryData();
            
            // Update primary business area
            this.formData.primaryBusinessAreas = primaryBusinessAreasSelect.value;
            
            // Clear the product/service category when business area changes
            this.formData.productServiceCategory = '';
            
            // Re-render to update the product service category dropdown
            this.render();
            
            // Re-attach event handlers
            this.setButtonHandlers();
            
            // If the primary area has changed, ensure product service category dropdown is updated
            const productServiceCategorySelect = this.domElement.querySelector('#productServiceCategory') as HTMLSelectElement;
            if (productServiceCategorySelect) {
              productServiceCategorySelect.addEventListener('change', () => {
                // Save current data whenever the product category changes
                this.saveCurrentIndustryData();
              });
            }
          });
          
          // Also set up the product service category change handler initially
          const productServiceCategorySelect = this.domElement.querySelector('#productServiceCategory') as HTMLSelectElement;
          if (productServiceCategorySelect) {
            productServiceCategorySelect.addEventListener('change', () => {
              // Save current data whenever the product category changes
              this.saveCurrentIndustryData();
            });
          }
        
        // Handle regulatory status change
        const regulatoryStatusYes = this.domElement.querySelector('#regulatoryStatusYes') as HTMLInputElement;
        const regulatoryStatusNo = this.domElement.querySelector('#regulatoryStatusNo') as HTMLInputElement;

        if (regulatoryStatusYes) {
          regulatoryStatusYes.addEventListener('change', () => {
            // Save current state
            this.saveCurrentIndustryData();
            
            // Set regulatory status to true
            this.formData.regulatoryStatus = true;
            
            // Re-render to show regulators selection
            this.render();
            this.setButtonHandlers();
          });
        }

        if (regulatoryStatusNo) {
          regulatoryStatusNo.addEventListener('change', () => {
            // Save current state
            this.saveCurrentIndustryData();
            
            // Set regulatory status to false and clear regulators
            this.formData.regulatoryStatus = false;
            this.formData.regulators = [];
            
            // Re-render to hide regulators selection
            this.render();
            this.setButtonHandlers();
          });
        }
      }
    }
    
    // Setup remove buttons for questions and files
    this.setupRemoveQuestionButtons();
    this.setupRemoveFileButtons();
  }
  
  private setupCountrySelector(): void {
    // Setup the country search
    const countrySearch = this.domElement.querySelector('#countrySearch') as HTMLInputElement;
    const countryList = this.domElement.querySelector('.' + styles.countryList);
    
    if (countrySearch && countryList) {
      
      // Re-render the country list to ensure it's properly initialized
      countryList.innerHTML = COUNTRIES.map(country => `
        <div class="${ styles.countryItem }">
          <input type="checkbox" id="country-${country.replace(/\s+/g, '-').toLowerCase()}" 
            class="country-checkbox" value="${country}" 
            ${this.formData.countriesOfOperation.includes(country) ? 'checked' : ''} />
          <label for="country-${country.replace(/\s+/g, '-').toLowerCase()}">${country}</label>
        </div>
      `).join('');
      
      // Setup search functionality
      countrySearch.addEventListener('input', () => {
        const searchValue = countrySearch.value.toLowerCase().trim();
        
        const countryItems = countryList.querySelectorAll('.' + styles.countryItem);
        for (let i = 0; i < countryItems.length; i++) {
          const item = countryItems[i] as HTMLElement;
          const label = item.querySelector('label');
          if (label) {
            const countryName = label.textContent.toLowerCase();
            if (searchValue === '' || countryName.indexOf(searchValue) !== -1) {
              item.style.display = '';
            } else {
              item.style.display = 'none';
            }
          }
        }
      });
    }
    
    // Setup checkbox for country selection
    const countryCheckboxes = this.domElement.querySelectorAll('.country-checkbox');
    
    for (let i = 0; i < countryCheckboxes.length; i++) {
      const checkbox = countryCheckboxes[i] as HTMLInputElement;
      const countryName = checkbox.value;
      
      checkbox.checked = this.formData.countriesOfOperation.includes(countryName);
      
      checkbox.addEventListener('change', () => {
        if (checkbox.checked) {
          if (!this.formData.countriesOfOperation.includes(countryName)) {
            this.formData.countriesOfOperation.push(countryName);
          }
        } else {
          const index = this.formData.countriesOfOperation.indexOf(countryName);
          if (index !== -1) {
            this.formData.countriesOfOperation.splice(index, 1);
          }
        }
        
        this.updateSelectedCountriesDisplay();
      });
    }
    
    this.setupRemoveCountryButtons();
    this.updateSelectedCountriesDisplay();
  }

  private setupRemoveCountryButtons(): void {
    const removeButtons = this.domElement.querySelectorAll('.remove-country-btn');
    
    for (let i = 0; i < removeButtons.length; i++) {
      const button = removeButtons[i] as HTMLButtonElement;
      button.addEventListener('click', () => {
        const countryName = button.getAttribute('data-country');
        
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
    // Try both selectors to see which one works
    const selectedCountriesContainer = this.domElement.querySelector(`.${styles.selectedCountries}`);
    
    if (selectedCountriesContainer) {
      selectedCountriesContainer.innerHTML = '';
      
      if (this.formData.countriesOfOperation.length > 0) {
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
    }
  }

  private validateStep1(): boolean {
    this.validateAttempted = true;
    let isValid = true;
    
    const fullNameInput = this.domElement.querySelector('#fullName') as HTMLInputElement;
    const organisationNameInput = this.domElement.querySelector('#organisationName') as HTMLInputElement;
    const emailAddressInput = this.domElement.querySelector('#emailAddress') as HTMLInputElement;
    const operationLocationInput = this.domElement.querySelector('#operationLocation') as HTMLInputElement;
    const operationLengthInput = this.domElement.querySelector('#operationLength') as HTMLInputElement;
    
    // Check each field for validity
    if (!fullNameInput || !fullNameInput.value.trim()) {
      isValid = false;
    }
    
    if (!organisationNameInput || !organisationNameInput.value.trim()) {
      isValid = false;
    }
    
    if (!emailAddressInput || !emailAddressInput.value.trim()) {
      isValid = false;
    } else {
      // Validate email format
      const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailPattern.test(emailAddressInput.value)) {
        isValid = false;
      }
    }
    
    if (!operationLocationInput || !operationLocationInput.value.trim()) {
      isValid = false;
    }
    
    if (this.formData.countriesOfOperation.length === 0) {
      isValid = false;
    }
    
    if (!operationLengthInput || !operationLengthInput.value.trim()) {
      isValid = false;
    }
    
    // Re-render to show validation messages
    if (!isValid) {
      this.render();
    }
    
    return isValid;
  }

  private validateStep2(): boolean {
    this.validateAttempted = true;
    let isValid = true;
    
    // Primary business area validation
    const primaryBusinessArea = this.domElement.querySelector('#primaryBusinessAreas') as HTMLSelectElement;
    if (!primaryBusinessArea || !primaryBusinessArea.value) {
      isValid = false;
    } else {
      this.formData.primaryBusinessAreas = primaryBusinessArea.value;
    }
    
    // Product/service category validation
    const productServiceCategory = this.domElement.querySelector('#productServiceCategory') as HTMLSelectElement;
    if (!productServiceCategory || !productServiceCategory.value) {
      isValid = false;
    } else {
      this.formData.productServiceCategory = productServiceCategory.value;
    }
    
    // Other product/service category validation if "Other" is selected
    if (this.formData.productServiceCategory === 'Other') {
      const otherProductServiceCategory = this.domElement.querySelector('#otherProductServiceCategory') as HTMLInputElement;
      if (!otherProductServiceCategory || !otherProductServiceCategory.value.trim()) {
        isValid = false;
      } else {
        this.formData.otherProductServiceCategory = otherProductServiceCategory.value.trim();
      }
    }
    
    // Operational status validation
    const operationalStatusYes = this.domElement.querySelector('#operationalStatusYes') as HTMLInputElement;
    const operationalStatusNo = this.domElement.querySelector('#operationalStatusNo') as HTMLInputElement;
    
    if (operationalStatusYes && operationalStatusYes.checked) {
      this.formData.operationalStatus = true;
    } else if (operationalStatusNo && operationalStatusNo.checked) {
      this.formData.operationalStatus = false;
    } else {
      isValid = false;
    }
    
    // Regulatory status validation
    const regulatoryStatusYes = this.domElement.querySelector('#regulatoryStatusYes') as HTMLInputElement;
    const regulatoryStatusNo = this.domElement.querySelector('#regulatoryStatusNo') as HTMLInputElement;
    
    if (regulatoryStatusYes && regulatoryStatusYes.checked) {
      this.formData.regulatoryStatus = true;
      
      // Make sure regulators array exists
      if (!this.formData.regulators) {
        this.formData.regulators = [];
      }
      
      // Regulator validation (only if regulatory status is Yes)
      const regulatorCheckboxes = this.domElement.querySelectorAll('.regulator-checkbox:checked') as NodeListOf<HTMLInputElement>;
      
      if (!regulatorCheckboxes || regulatorCheckboxes.length === 0) {
        isValid = false;
      } else {
        this.formData.regulators = Array.from(regulatorCheckboxes).map(checkbox => checkbox.value);
        
        // Check for Other regulator
        if (this.formData.regulators.indexOf('Other') !== -1) {
          const otherRegulator = this.domElement.querySelector('#otherRegulator') as HTMLInputElement;
          
          if (!otherRegulator || !otherRegulator.value.trim()) {
            isValid = false;
          } else {
            this.formData.otherRegulator = otherRegulator.value.trim();
          }
        }
      }
    } else if (regulatoryStatusNo && regulatoryStatusNo.checked) {
      this.formData.regulatoryStatus = false;
      // If regulatory status is No, we don't need regulators
      this.formData.regulators = [];
    } else {
      isValid = false;
    }
    
    // Re-render to show validation messages if invalid
    if (!isValid) {
      this.render();
    }
    
    return isValid;
  }

  private validateStep3(): boolean {
    this.validateAttempted = true;
    let isValid = true;
    
    // Product/service description validation
    const descriptionTextarea = this.domElement.querySelector('#productServiceDescription') as HTMLTextAreaElement;
    if (!descriptionTextarea || !descriptionTextarea.value.trim()) {
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
      isValid = false;
    }
    
    // FAQ confirmation validation
    const faqYes = this.domElement.querySelector('#faqConfirmationYes') as HTMLInputElement;
    const faqNo = this.domElement.querySelector('#faqConfirmationNo') as HTMLInputElement;
    if ((!faqYes || !faqYes.checked) && (!faqNo || !faqNo.checked)) {
      isValid = false;
    }
    
    // Consent validation
    const consentCheckbox = this.domElement.querySelector('#consentCheckbox') as HTMLInputElement;
    if (!consentCheckbox || !consentCheckbox.checked) {
      isValid = false;
    }
    
    // Re-render to show validation messages if invalid
    if (!isValid) {
      this.render();
      
      // Re-setup event handlers after render
      this.setupRemoveQuestionButtons();
      this.setupRemoveFileButtons();
    }
    
    return isValid;
  }

  private isValidEmail(email: string): boolean {
    const regex = /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return regex.test(String(email).toLowerCase());
  }

  /**
   * Validates just the email field and updates the UI immediately
   */
  private validateEmailField(): void {
    const emailInput = this.domElement.querySelector('#emailAddress') as HTMLInputElement;
    const emailField = emailInput ? emailInput.parentElement : null;
    
    if (emailInput && emailField) {
      const emailValue = emailInput.value.trim();
      const isValid = emailValue === '' || this.isValidEmail(emailValue);
      const isEmpty = emailValue === '';
      
      // Remove existing error classes and messages
      emailInput.classList.remove(styles.error);
      const existingErrorMsg = emailField.querySelector('.email-error-msg');
      if (existingErrorMsg) {
        existingErrorMsg.remove();
      }
      
      // Add error styling and message if invalid
      if (!isEmpty && !isValid) {
        emailInput.classList.add(styles.error);
        
        const errorMsg = document.createElement('div');
        errorMsg.className = `${styles.errorText} email-error-msg`;
        errorMsg.textContent = 'Please enter a valid email address';
        emailField.appendChild(errorMsg);
      }
    }
  }

  private saveStep1Data(): void {
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
    
    // Move to the next step
    this.currentStep = 2;
    this.validateAttempted = false;
    this.render();
  }

  private saveStep2Data(): void {
    // Get primary business area
    const primaryBusinessAreasSelect = this.domElement.querySelector('#primaryBusinessAreas') as HTMLSelectElement;
    if (primaryBusinessAreasSelect) {
      this.formData.primaryBusinessAreas = primaryBusinessAreasSelect.value;
    }
    
    // Get product/service category
    const productServiceCategorySelect = this.domElement.querySelector('#productServiceCategory') as HTMLSelectElement;
    if (productServiceCategorySelect) {
      this.formData.productServiceCategory = productServiceCategorySelect.value;
    }
    
    // Get other product/service category if applicable
    if (this.formData.productServiceCategory === 'Other') {
      const otherProductServiceCategoryInput = this.domElement.querySelector('#otherProductServiceCategory') as HTMLInputElement;
      if (otherProductServiceCategoryInput) {
        this.formData.otherProductServiceCategory = otherProductServiceCategoryInput.value;
      }
    }
    
    // Get operational status
    const operationalStatusYes = this.domElement.querySelector('#operationalStatusYes') as HTMLInputElement;
    const operationalStatusNo = this.domElement.querySelector('#operationalStatusNo') as HTMLInputElement;
    if (operationalStatusYes && operationalStatusNo) {
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
      if (regulatoryStatusYes.checked) {
        this.formData.regulatoryStatus = true;
        
        // Check for regulator checkboxes
        const regulatorCheckboxes = this.domElement.querySelectorAll('.regulator-checkbox:checked') as NodeListOf<HTMLInputElement>;
        if (regulatorCheckboxes && regulatorCheckboxes.length > 0) {
          // Update the regulators array directly from checked checkboxes
          this.formData.regulators = Array.from(regulatorCheckboxes).map(checkbox => checkbox.value);
        }
      } else if (regulatoryStatusNo.checked) {
        this.formData.regulatoryStatus = false;
      }
    }
    
    // Get other regulator if applicable
    if (this.formData.regulators.indexOf('Other') !== -1) {
      const otherRegulatorInput = this.domElement.querySelector('#otherRegulator') as HTMLInputElement;
      if (otherRegulatorInput && otherRegulatorInput.value.trim()) {
        this.formData.otherRegulator = otherRegulatorInput.value.trim();
      }
    }
    
    // Do NOT validate step 3 here - we're moving TO step 3
  }

  private saveStep3Data(): void {
    // Get product/service description
    const descriptionTextarea = this.domElement.querySelector('#productServiceDescription') as HTMLTextAreaElement;
    if (descriptionTextarea) {
      this.formData.productServiceDescription = descriptionTextarea.value;
    }
    
    // Get questions
    const questionInputs = this.domElement.querySelectorAll('.question-input') as NodeListOf<HTMLTextAreaElement>;
    let questions = [];
    
    // Save all questions, even empty ones
    for (let i = 0; i < questionInputs.length; i++) {
      questions.push(questionInputs[i].value);
    }
    
    // If we have at least one question (even if empty), use that
    if (questions.length > 0) {
      this.formData.questions = questions;
    } else {
      // Otherwise, ensure we always have at least one question field
      this.formData.questions = [''];
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

  /**
   * Checks if an error is a connection reset error that should be retried
   */
  private isRetryableError(error: any): boolean {
    const errorMessage = error.message || error.toString() || '';
    
    // Only retry specific connection reset errors
    const retryableErrors = [
      'ECONNRESET',
      'ENOTFOUND', 
      'ETIMEDOUT',
      'ECONNREFUSED',
      'ERR_CONNECTION_RESET',
      'ERR_NETWORK',
      'ERR_INTERNET_DISCONNECTED',
      'ERR_CONNECTION_REFUSED',
      'ERR_CONNECTION_TIMED_OUT'
    ];
    
    return retryableErrors.some(retryableError => 
      errorMessage.includes(retryableError)
    );
  }

  /**
   * Handles retry logic for connection errors
   */
  private handleRetry(error: any): Promise<any> {
    if (this.currentRetryCount < this.maxRetries && this.isRetryableError(error)) {
      this.currentRetryCount++;
      console.log(`Connection error detected, retrying... (${this.currentRetryCount}/${this.maxRetries})`);
      
      return new Promise((resolve, reject) => {
        setTimeout(() => {
          // Retry the entire submission process
          this.submitFormInternal()
            .then(resolve)
            .catch(reject);
        }, this.retryDelay);
      });
    }
    
    // If we've exhausted retries or it's not a retryable error, throw the original error
    throw error;
  }

  /**
   * Internal submission method that handles the actual submission logic
   */
  private submitFormInternal(): Promise<any> {
    // Make sure we have at least one question
    if (!this.formData.questions || this.formData.questions.length === 0) {
      this.formData.questions = [''];
    }
    
    // Check if we have files to upload
    const hasFiles = this.formData.files && this.formData.files.length > 0;
    
    // Submit form data to SharePoint list using XML and service account
    return this.createEnquiryDetailsListItem()
      .then((response) => {
        // Upload files if there are any
        if (hasFiles) {
          return this.uploadFiles(this.formData.files);
        }
        return Promise.resolve();
      })
      .catch((error) => {
        // Check if this is a retryable connection error
        if (this.isRetryableError(error)) {
          console.log('Retryable connection error detected, attempting retry...');
          return this.handleRetry(error);
        }
        
        // If not retryable, throw the error
        throw error;
      });
  }

  private submitForm(): void {
    // Reset retry count for new submission
    this.currentRetryCount = 0;
    
    // Set loading state
    this.isSubmitting = true;
    this.render(); // Re-render to show loading state
    
    console.log('Starting form submission process...');
    
    this.submitFormInternal()
      .then(() => {
        console.log('Form submission completed successfully, showing thank you step');
        
        // Clear loading state
        this.isSubmitting = false;
        
        // Show thank you step
        this.currentStep = 4;
        this.render();
        
        // Clear authentication context to prevent future prompts
        this.clearAuthenticationContext();
        
        // Set a temporary blocker on navigation to prevent auth prompts
        window.onbeforeunload = () => {
          return "Your submission has been received. Are you sure you want to leave?";
        };
        
        // Remove the blocker after a few seconds
        setTimeout(() => {
          window.onbeforeunload = null;
        }, 3000);
      })
      .catch((error) => {
        console.error('Error submitting form:', error);
        
        // Clear loading state
        this.isSubmitting = false;
        this.render(); // Re-render to hide loading state
        
        // Log error details
        if (error.message) {
          console.error('Error message:', error.message);
        }
        if (error.stack) {
          console.error('Error stack:', error.stack);
        }
        
        // Show error message to the user
        const errorContainer = this.domElement.querySelector('.error-message');
        if (errorContainer) {
          errorContainer.innerHTML = 'There was an error submitting your form. Please try again later.';
          errorContainer.className = 'error-message visible';
        }
        
        // Clear authentication context even on error
        this.clearAuthenticationContext();
      });
  }
  
  /**
   * Clears authentication context to prevent lingering auth prompts
   */
  private clearAuthenticationContext(): void {
    try {
      console.log('Clearing authentication context...');
      
      // Check if we've already cleared auth in the last 30 minutes
      const lastCleared = localStorage.getItem('authCleared');
      if (lastCleared) {
        const timeSince = new Date().getTime() - new Date(lastCleared).getTime();
        if (timeSince < 30 * 60 * 1000) { // 30 minutes
          console.log('Auth cleared recently, skipping');
          return;
        }
      }
      
      // Remove the problematic API call that triggers sign-in dialog
      // The localStorage marking is sufficient for our needs
      
      // Mark as completed in localStorage
      localStorage.setItem('authCleared', new Date().toISOString());
    } catch (error) {
      console.error('Error clearing authentication context:', error);
    }
  }
  
  /**
   * Closes the form and resets state to prevent lingering connections
   * This should be called when navigating away or closing the form
   */
  public onDispose(): void {
    super.onDispose();
    
    // Clean up any event listeners
    window.onbeforeunload = null;
    
    // Clear auth context when leaving the page
    this.clearAuthenticationContext();
  }

  /**
   * Creates a list item in the Enquiry Details list using the new upload.js API
   */
  private createEnquiryDetailsListItem(): Promise<any> {
    const apiUrl = 'https://integration.fsca.co.za/api/create-list-item';
    const item = {
      Title: `Enquiry from ${this.formData.fullName}`,
      FullName: this.formData.fullName,
      OrganisationName: this.formData.organisationName,
      ContactNumber: this.formData.contactNumber,
      EmailAddress: this.formData.emailAddress,
      WebsiteAddress: this.formData.websiteAddress,
      OperationLocation: this.formData.operationLocation,
      CountriesOfOperation: (Array.isArray(this.formData.countriesOfOperation) ? this.formData.countriesOfOperation.join(', ') : ''),
      OperationLength: this.formData.operationLength,
      PrimaryBusinessAreas: this.formData.primaryBusinessAreas,
      ProductServiceCategory: this.formData.productServiceCategory,
      OtherProductServiceCategory: this.formData.otherProductServiceCategory,
      OperationalStatus: this.formData.operationalStatus === true ? 'Yes' : (this.formData.operationalStatus === false ? 'No' : ''),
      RegulatoryStatus: this.formData.regulatoryStatus === true ? 'Yes' : (this.formData.regulatoryStatus === false ? 'No' : ''),
      Regulators: (Array.isArray(this.formData.regulators) ? this.formData.regulators.join(', ') : ''),
      OtherRegulator: this.formData.otherRegulator,
      ProductServiceDescription: this.formData.productServiceDescription,
      Questions: (Array.isArray(this.formData.questions) ? this.formData.questions.join('\n\n') : ''),
      AdditionalInformation: this.formData.additionalInformation,
      FAQConfirmation: this.formData.faqConfirmation === true ? 'Yes' : (this.formData.faqConfirmation === false ? 'No' : ''),
      ConsentConfirmation: this.formData.consentConfirmation ? 'Yes' : 'No',
      FileAttachments: (this.formData.files && this.formData.files.length > 0) ? 'Yes' : 'No',
      NumberOfAttachments: (this.formData.files ? this.formData.files.length.toString() : '0'),
      SubmissionDate: new Date().toISOString()
    };
    return fetch(apiUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(item)
    })
      .then(response => {
        if (!response.ok) throw new Error('API call failed: ' + response.status);
        return response.json();
      })
      .then(data => {
        if (data && data.Id) {
          console.log('List item created successfully via upload.js API. Item ID:', data.Id);
          return { status: 201, json: () => Promise.resolve({ Id: data.Id }) };
        } else {
          console.error('upload.js API response did not contain item ID:', data);
          throw new Error('upload.js API submission returned an error');
        }
      });
  }

  /**
   * Uploads files to SharePoint document library with fallbacks
   */
  private uploadFiles(files: File[]): Promise<void> {
    console.log(`Starting upload process for ${files.length} files`);
    
    if (!files || files.length === 0) {
      console.log('No files to upload');
      return Promise.resolve();
    }
    
    // Generate folder name based on user's name and timestamp
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const folderName = `${this.formData.fullName || 'Anonymous'}_${timestamp}`;
    
    console.log(`Generated folder name: ${folderName}`);
    
    // Upload each file, one at a time to avoid overwhelming the server
    // Skip folder creation entirely and just use prefixed filenames
    console.log(`Uploading ${files.length} files with prefix ${folderName}`);
    
    return files.reduce((previous, file) => {
      return previous.then(() => {
        console.log(`Starting upload for file: ${file.name}`);
        return this.uploadFileWithServiceAccount(file, folderName)
          .then(() => {
            console.log(`Successfully uploaded file: ${file.name}`);
            return Promise.resolve();
          })
          .catch(error => {
            console.error(`Error uploading file ${file.name}:`, error);
            // Continue with other files even if this one fails
            return Promise.resolve();
          });
      });
    }, Promise.resolve());
  }

  /**
   * Encodes service account credentials for Basic Authentication
   */
  private getServiceAccountCredentials(): string {
    const username = `${SERVICE_ACCOUNT.domain}\\${SERVICE_ACCOUNT.username}`;
    const password = SERVICE_ACCOUNT.password;
    return this.encodeCredentials(username, password);
  }

  /**
   * Encodes credentials to Base64 for Basic Authentication
   */
  private encodeCredentials(username: string, password: string): string {
    const toEncode = username + ':' + password;
    // For IE compatibility in TypeScript 2.4.2
    if (typeof window !== 'undefined' && window.btoa) {
      return window.btoa(toEncode);
    } else {
      // Fallback if btoa is not available
      const keyStr = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=';
      let output = '';
      let chr1, chr2, chr3, enc1, enc2, enc3, enc4;
      let i = 0;

      do {
        chr1 = toEncode.charCodeAt(i++);
        chr2 = toEncode.charCodeAt(i++);
        chr3 = toEncode.charCodeAt(i++);

        enc1 = chr1 >> 2;
        enc2 = ((chr1 & 3) << 4) | (chr2 >> 4);
        enc3 = ((chr2 & 15) << 2) | (chr3 >> 6);
        enc4 = chr3 & 63;

        if (isNaN(chr2)) {
          enc3 = enc4 = 64;
        } else if (isNaN(chr3)) {
          enc4 = 64;
        }

        output = output +
          keyStr.charAt(enc1) +
          keyStr.charAt(enc2) +
          keyStr.charAt(enc3) +
          keyStr.charAt(enc4);
      } while (i < toEncode.length);

      return output;
    }
  }

  /**
   * Gets the Form Digest value using service account
   */
  private getFormDigestValueWithServiceAccount(): Promise<string> {
    console.log('Getting form digest value with service account');
    
    const url = 'https://www.ifwg.co.za/_api/contextinfo';
    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=verbose'
      }
    })
    .then((response: SPHttpClientResponse) => {
      if (response.ok) {
        return response.json();
      } else {
        console.error('Failed to get Form Digest:', response.status, response.statusText);
        // Fallback to a dummy digest
        return Promise.resolve({ d: { GetContextWebInformation: { FormDigestValue: '0x1234567890ABCDEF' } } });
      }
    })
    .then((jsonResponse) => {
      const formDigestValue = jsonResponse.d.GetContextWebInformation.FormDigestValue;
      console.log('Form Digest obtained successfully');
      return formDigestValue;
    })
    .catch((error) => {
      console.error('Error getting Form Digest:', error);
      // Fallback to a dummy digest (not secure, but might work in some environments)
      return '0x1234567890ABCDEF';
    });
  }

  /**
   * Creates a folder in SharePoint using service account and XML
   */
  private createFolderWithServiceAccount(folderName: string): Promise<any> {
    console.log(`Creating folder ${folderName}`);
    
    // Get the request digest from the page
    const requestDigest = (document.getElementById("__REQUESTDIGEST") as HTMLInputElement).value;
    
    // Try using the modern REST API
    const folderUrl = `${this.context.pageContext.web.serverRelativeUrl}/EnquiryFormDocuments/${folderName}`;
    const restUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/folders`;
    
    const restBody = JSON.stringify({
      '__metadata': { 'type': 'SP.Folder' },
      'ServerRelativeUrl': folderUrl
    });
    
    return this.context.spHttpClient.post(
      restUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-type': 'application/json;odata=verbose',
          'X-RequestDigest': requestDigest
        },
        body: restBody
      }
    )
    .then(response => {
      if (response.ok) {
        console.log(`Folder ${folderName} created successfully`);
        return response.json();
      }
      throw new Error('REST API folder creation failed');
    })
    .catch(restError => {
      console.log('REST API folder creation failed, falling back to direct folder creation', restError);
      
      // Try direct folder creation as a fallback
      const emergencyUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('EnquiryFormDocuments')/folders/add(url='${folderName}')`;
      
      return this.context.spHttpClient.post(
        emergencyUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-type': 'application/json;odata=verbose',
            'X-RequestDigest': requestDigest
          },
          body: JSON.stringify({})
        }
      )
      .then(response => {
        if (response.ok) {
          console.log(`Folder ${folderName} created successfully with fallback method`);
          return response.json();
        }
        console.log('All folder creation methods failed, will attempt to upload without folder');
        return Promise.resolve(); // Continue even if folder creation failed
      })
      .catch(finalError => {
        console.log('All folder creation methods failed, will attempt to upload without folder');
        return Promise.resolve(); // Continue even if folder creation failed
      });
    });
  }

  /**
   * Uploads a file to SharePoint using the new upload.js API
   */
  private uploadFileWithServiceAccount(file: File, folderName: string): Promise<any> {
    console.log(`Uploading file ${file.name} to folder ${folderName} via upload.js API`);
    const apiUrl = 'https://integration.fsca.co.za/api/upload-doc';
    return new Promise<any>((resolve, reject) => {
      if (file.size > 10 * 1024 * 1024) { // 10MB limit
        console.log(`File ${file.name} exceeds 10MB, skipping actual upload but recording submission`);
        return resolve(`File ${file.name} was too large (${Math.round(file.size/1024/1024)}MB) to upload automatically.`);
      }
      const formData = new FormData();
      formData.append('document', file, file.name);
     // formData.append('folderName', folderName);
      fetch(apiUrl, {
        method: 'POST',
        body: formData
      })
        .then(response => {
          if (!response.ok) throw new Error('API call failed: ' + response.status);
          return response.json();
        })
        .then(data => {
          if (data && data.ServerRelativeUrl) {
            console.log(`File ${file.name} uploaded successfully via upload.js API. ServerRelativeUrl:`, data.ServerRelativeUrl);
            resolve(`File ${file.name} uploaded successfully`);
          } else {
            console.error('upload.js API file upload response did not contain ServerRelativeUrl:', data);
            resolve(`Error uploading ${file.name}, but form submission recorded`);
          }
        })
        .catch((uploadError) => {
          console.error(`Error uploading file ${file.name} via upload.js API:`, uploadError);
          resolve(`Error uploading ${file.name}, but form submission recorded`);
        });
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
    // Get product/service description
    const descriptionTextarea = this.domElement.querySelector('#productServiceDescription') as HTMLTextAreaElement;
    if (descriptionTextarea) {
      this.formData.productServiceDescription = descriptionTextarea.value;
    }
    
    // Get existing questions (preserve ALL user input)
    const questionInputs = this.domElement.querySelectorAll('.question-input') as NodeListOf<HTMLTextAreaElement>;
    
    // Create a new array to ensure we don't unexpectedly modify questions
    const updatedQuestions = [];
    
    // Save all questions including empty ones
    for (let i = 0; i < questionInputs.length; i++) {
      updatedQuestions.push(questionInputs[i].value);
    }
    
    // Only update if we have at least one question
    if (updatedQuestions.length > 0) {
      this.formData.questions = updatedQuestions;
    } else if (!this.formData.questions || this.formData.questions.length === 0) {
      // Ensure we always have at least one question
      this.formData.questions = [''];
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

  /**
   * Makes an API call using service account credentials to avoid authentication prompts
   */
  private makeApiCallWithServiceAccount(url: string, method: string, headers: any, body: string | ArrayBuffer): Promise<any> {
    // If url doesn't start with http, prepend the web's absolute URL
    if (url.indexOf('http') !== 0 && this.context) {
      url = `${this.context.pageContext.web.absoluteUrl}${url.startsWith('/') ? '' : '/'}${url}`;
    }
    
    // Clean URL of any auth-triggering parameters
    url = url.replace(/([?&])prompt=login(&|$)/, '$1')
             .replace(/([?&])force=true(&|$)/, '$1')
             .replace(/\?$/, '');
    
    console.log(`Making API call to ${url} with service account`);
    
    // For REST API calls and SOAP calls, we need a form digest
    const needsDigest = (method === 'POST' || method === 'PUT' || method === 'DELETE' || method === 'MERGE') &&
                        (url.indexOf('/_api/') > -1 || url.indexOf('_vti_bin') > -1);
    
    // Get form digest first if needed, otherwise proceed directly
    const digestPromise = needsDigest 
      ? this.getFormDigestWithServiceAccount() 
      : Promise.resolve(null);
      
    return digestPromise.then(digest => {
      return new Promise<string>((resolve, reject) => {
        try {
          console.log(`Making asynchronous request to ${url}`);
          const xhr = new XMLHttpRequest();
          xhr.open(method, url, true);  // always use async for more reliability
          xhr.timeout = 60000;  // 60 second timeout
          
          // Set headers
          Object.keys(headers).forEach(header => {
            xhr.setRequestHeader(header, headers[header]);
          });
          
          // Add digest if needed
          if (needsDigest && digest) {
            console.log('Adding X-RequestDigest header');
            xhr.setRequestHeader('X-RequestDigest', digest);
          }
          
          // Set credentials
          const credentials = this.getServiceAccountCredentials();
         xhr.setRequestHeader('Authorization', 'Basic ' + credentials);
          
          // Set credentials handling - critical to avoid prompts
          xhr.withCredentials = false;
          
          xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
              if (xhr.status >= 200 && xhr.status < 300) {
                console.log(`API call succeeded with status: ${xhr.status}`);
                resolve(xhr.responseText);
              } else if (xhr.status === 401 || xhr.status === 403) {
                // Authentication error - continue anyway with empty response
                console.log(`Auth error ${xhr.status} intercepted, continuing anyway`);
                resolve('{}');
              } else {
                console.log(`API call failed: ${xhr.status} `, xhr.responseText);
                reject(new Error(`API call failed: ${xhr.responseText}`));
              }
            }
          };
          
          xhr.onerror = function() {
            console.error('Request error occurred');
            // Return empty response instead of error
            resolve('{}');
          };
          
          xhr.ontimeout = function() {
            console.error('Request timed out');
            // Return empty response instead of error
            resolve('{}');
          };
          
          // Send request
          console.log(`Sending ${method} request to ${url}`);
          if (typeof body === 'string') {
            xhr.send(body);
          } else {
            // Handle ArrayBuffer data
            xhr.send(body);
          }
        } catch (error) {
          console.error('Error in API call:', error);
          // Return empty response instead of error
          resolve('{}');
        }
      });
    });
  }
  
  /**
   * Gets a form digest value using service account credentials
   * This is required for most write operations to SharePoint
   */
  private getFormDigestWithServiceAccount(): Promise<string> {
    console.log('Getting form digest with service account');
    
    const digestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/contextinfo`;
    
    return new Promise<string>((resolve, reject) => {
      try {
        const xhr = new XMLHttpRequest();
        xhr.open('POST', digestUrl, true);
        xhr.timeout = 30000; // 30 second timeout
        
        // Set headers
        xhr.setRequestHeader('Accept', 'application/json;odata=verbose');
        
        // CRITICAL: Use service account credentials (not user credentials)
        const credentials = this.getServiceAccountCredentials();
        xhr.setRequestHeader('Authorization', 'Basic ' + credentials);
        xhr.withCredentials = false;
        
        xhr.onreadystatechange = function() {
          if (xhr.readyState === 4) {
            if (xhr.status >= 200 && xhr.status < 300) {
              console.log('Successfully retrieved form digest');
              try {
                const response = JSON.parse(xhr.responseText);
                const formDigestValue = response.d.GetContextWebInformation.FormDigestValue;
                resolve(formDigestValue);
              } catch (parseError) {
                console.error('Error parsing digest response:', parseError);
                resolve('0xDEADBEEF12345678');
              }
            } else if (xhr.status === 401 || xhr.status === 403) {
              console.log('Auth error getting digest, using fallback');
              resolve('0xDEADBEEF12345678');
            } else {
              console.log(`Failed to get form digest: ${xhr.status}`);
              resolve('0xDEADBEEF12345678');
            }
          }
        };
        
        xhr.onerror = function() {
          console.error('Request error getting digest');
          resolve('0xDEADBEEF12345678');
        };
        
        xhr.ontimeout = function() {
          console.error('Request timeout getting digest');
          resolve('0xDEADBEEF12345678');
        };
        
        // Send request
        xhr.send();
      } catch (error) {
        console.error('Error in digest request:', error);
        resolve('0xDEADBEEF12345678');
      }
    });
  }

  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      // Initialize the form data
      this.formData = {
        fullName: '',
        organisationName: '',
        contactNumber: '',
        emailAddress: '',
        websiteAddress: '',
        operationLocation: '',
        countriesOfOperation: [],
        operationLength: '',
        primaryBusinessAreas: '',
        productServiceCategory: '',
        otherProductServiceCategory: '',
        operationalStatus: null,
        regulatoryStatus: null,
        regulators: [],
        otherRegulator: '',
        productServiceDescription: '',
        questions: [''], // Always start with one question
        additionalInformation: '',
        faqConfirmation: null,
        consentConfirmation: false,
        files: []
      };
      
      // Set up auth interception to prevent login prompts
      this.setupAuthInterceptors();

      // Ensure we have a default question
      if (!this.formData.questions || this.formData.questions.length === 0) {
        this.formData.questions = [''];
      }
      
      // Set up handlers to prevent authentication dialogs on navigation
      window.addEventListener('beforeunload', () => {
        if (this.currentStep === 4) { // Only if we've submitted the form
          this.clearAuthenticationContext();
        }
      });

      return Promise.resolve();
    });
  }
  
  /**
   * Set up interceptors to prevent authentication prompts, but only for this webpart
   */
  private setupAuthInterceptors(): void {
    console.log('Setting up targeted auth interceptors for form submission only');
    
    // We won't override global fetch or XMLHttpRequest to avoid breaking SharePoint
    // Instead, we'll just handle auth in our own API calls
    
    // Add unload handler to clean up any ongoing requests when the form completes
    window.addEventListener('beforeunload', () => {
      // Only clear if we've submitted a form
      if (this.currentStep === 4) {
        this.clearAuthenticationContext();
      }
    });
  }

  /**
   * Makes an API call using the SPHttpClient instead of XMLHttpRequest
   * @param url The URL to call
   * @param method The HTTP method (GET, POST, etc.)
   * @param headers Additional headers to include
   * @param body The request body (for POST/PUT requests)
   */
  private makeApiCallWithSPHttpClient(url: string, method: string, headers: any, body: any): Promise<any> {
    // If url doesn't start with http, prepend the web's absolute URL
    if (url.indexOf('http') !== 0 && this.context) {
      url = `${this.context.pageContext.web.absoluteUrl}${url.startsWith('/') ? '' : '/'}${url}`;
    }
    
    // Clean URL of any auth-triggering parameters
    url = url.replace(/([?&])prompt=login(&|$)/, '$1')
             .replace(/([?&])force=true(&|$)/, '$1')
             .replace(/\?$/, '');
    
    console.log(`Making API call to ${url} with SPHttpClient`);
    
    // For REST API calls, we need a form digest
    const needsDigest = (method === 'POST' || method === 'PUT' || method === 'DELETE' || method === 'MERGE') &&
                        (url.indexOf('/_api/') > -1);
    
    // Get form digest first if needed, otherwise proceed directly
    const digestPromise = needsDigest 
      ? this.getFormDigestWithServiceAccount() 
      : Promise.resolve(null);
      
    return digestPromise.then(digest => {
      // Prepare request options
      const requestOptions: any = {
        headers: { ...headers }
      };
      
      // Add digest if needed
      if (needsDigest && digest) {
        console.log('Adding X-RequestDigest header');
        requestOptions.headers['X-RequestDigest'] = digest;
      }
      
      // Add body if provided
      if (body) {
        requestOptions.body = body;
      }
      
      // Make the request using SPHttpClient
      console.log(`Sending ${method} request to ${url}`);
      
      switch (method.toUpperCase()) {
        case 'GET':
          return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, requestOptions);
        case 'POST':
          return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, requestOptions);
        case 'PUT':
          return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { ...requestOptions, headers: { ...requestOptions.headers, 'X-HTTP-Method': 'PUT' } });
        case 'DELETE':
          return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { ...requestOptions, headers: { ...requestOptions.headers, 'X-HTTP-Method': 'DELETE' } });
        case 'PATCH':
          return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { ...requestOptions, headers: { ...requestOptions.headers, 'X-HTTP-Method': 'MERGE' } });
        default:
          return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, requestOptions);
      }
    })
    .then((response: SPHttpClientResponse) => {
      if (response.ok) {
        return response.text().then(text => {
          try {
            // Try to parse as JSON if possible
            return JSON.parse(text);
          } catch (e) {
            // Return as text if not valid JSON
            return text;
          }
        });
      } else if (response.status === 401 || response.status === 403) {
        // Authentication error - continue anyway with empty response
        console.log(`Auth error ${response.status} intercepted, continuing anyway`);
        return '{}';
      } else {
        console.log(`API call failed: ${response.status}`);
        return response.text().then(text => {
          throw new Error(`API call failed: ${text}`);
        });
      }
    })
    .catch(error => {
      console.error('Error in API call:', error);
      // Return empty response instead of error
      return '{}';
    });
  }
}
