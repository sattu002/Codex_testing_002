// =====================================================================
// === THIS IMPORT BLOCK IS THE MOST CRITICAL PART OF THE FIX ===
// =====================================================================
import {
  db,
  dataCollection,
  doc,
  serverTimestamp,
  writeBatch,
} from "./firebase.js";
// =====================================================================


// --- Global Variables & Constants ---
window.SpeechRecognition =
  window.SpeechRecognition || window.webkitSpeechRecognition;
let recognition;
let currentTargetInput = null;
let collectedData = []; // Holds the customer entries for the CURRENT session
let isRecordingAllActive = false;
// MODIFIED: Added 'region' to the voice recording sequence
const recordAllSequence = [
  "branchName",
  "region",
  "customerName",
  "customerID",
  "mobileNumber",
  "savingsAccountNumber",
  "csbCode",
  "nomineeName",
  "nomineeRelationship",
  "nomineeMobileNumber",
];
let currentRecordAllIndex = -1;
let recordAllStatusElement = null;
let flatpickrInstances = []; // To store instances of the date picker

let partnerData = {
  BANGIYA: {
    Combo: [
      { Premium: 490, Tenure: 1, "CSE Name": "Jahed" },
      { Premium: 690, Tenure: 1, "CSE Name": "Jahed" },
      { Premium: 980, Tenure: 2, "CSE Name": "Jahed" },
      { Premium: 990, Tenure: 1, "CSE Name": "Jahed" },
    ],
    Telemedicine: [
      { Premium: 360, Tenure: 1, "CSE Name": "Jahed" },
      { Premium: 700, Tenure: 2, "CSE Name": "Jahed" },
      { Premium: 1000, Tenure: 3, "CSE Name": "Jahed" },
      { Premium: 2000, Tenure: 6, "CSE Name": "Jahed" },
      { Premium: 3000, Tenure: 9, "CSE Name": "Jahed" },
    ],
  },
  PBGB: {
    Combo: [
      { Premium: 490, Tenure: 1, "CSE Name": "Aditya" },
      { Premium: 690, Tenure: 1, "CSE Name": "Aditya" },
    ],
    Telemedicine: [{ Premium: 365, Tenure: 1, "CSE Name": "Aditya" }],
  },
  UBKGB: {
    Combo: [
      { Premium: 490, Tenure: 1, "CSE Name": "Abhijit" },
      { Premium: 690, Tenure: 1, "CSE Name": "Abhijit" },
    ],
    Telemedicine: [{ Premium: 365, Tenure: 1, "CSE Name": "Abhijit" }],
  },
  KCCB: {
    Combo: [
      { Premium: 700, Tenure: 1, "CSE Name": "Aditya" },
      { Premium: 1050, Tenure: 1, "CSE Name": "Aditya" },
    ],
    Telemedicine: [{ Premium: 399, Tenure: 1, "CSE Name": "Aditya" }],
  },
  "Assam Vikas Gramin Bank": {
    Telemedicine: [{ Premium: 365, Tenure: 1, "CSE Name": "Abhishek" }],
  },
  DCCB: {
    Telemedicine: [{ Premium: 365, Tenure: 1, "CSE Name": "Abhishek" }],
  },
  UBGB: {
    Telemedicine: [{ Premium: 365, Tenure: 1, "CSE Name": "Nazreen" }],
  },
};

const LS_DATA_KEY = "customerEntriesData_v8";
const LS_PARTNER_KEY = "customerPartnerData_v1";
const LS_THEME_KEY = "customerAppTheme";
const LS_UI_LANG_KEY = "customerAppUiLang_v2";
let genderChartInstance, branchChartInstance, partnerChartInstance;
let currentChartSlide = 0;
let uploadedImages = [];
let currentImageIndex = 0;
const MAX_IMAGE_UPLOADS = 50;
let isTableFullscreen = false;


// MODIFIED: Added "region" to the list of fields to be managed
const customerFieldKeys = [
  "branchName", "region", "customerName", "customerID", "gender", "mobileNumber", "enrolmentDate",
  "dob", "savingsAccountNumber", "csbCode", "d2cRoCode", "partnerName", "productName",
  "premium", "nomineeName", "nomineeRelationship", "nomineeDob", "nomineeMobileNumber",
  "nomineeGender", "remarks",
];

// MODIFIED: Added "region"
const formElementIds = [
  "partnerNameSelect", "productSelect", "premiumSelect", "branchName", "region", "customerName",
  "customerID", "mobileNumber", "enrolmentDate", "dob", "savingsAccountNumber",
  "csbCode", "d2cRoCode", "nomineeName", "nomineeRelationship", "nomineeDob",
  "nomineeMobileNumber", "remarks",
];

// MODIFIED: Added "region"
const textInputIds = formElementIds.filter(
  (k) => !k.includes("Select") && k !== "enrolmentDate" && k !== "dob" && k !== "nomineeDob" && k !== "remarks"
);

const optionalFields = [
  "branchName", "region", "customerName", "customerID", "mobileNumber", "enrolmentDate", "dob",
  "savingsAccountNumber", "csbCode", "d2cRoCode", "nomineeName", "nomineeRelationship",
  "nomineeDob", "nomineeMobileNumber", "remarks",
];

const duplicateCheckFields = ["customerName", "mobileNumber", "dob"];
const requiredFormElementIds = formElementIds.filter((f) => !optionalFields.includes(f));
const requiredRadioGroups = ["gender"];

const criticalFieldsForBlanks = [
  "customerName",
  "partnerName",
  "productName",
  "premium",
];

const ADMIN_USERNAME = "Satvick";
const ADMIN_PASSWORD = "satvickisbest";

// --- i18n Translations ---
const translations = {
  en: {
    // Toasts
    toastNothingToSubmit: "There is no data to submit.",
    toastSubmitSuccess: "{count} entries successfully submitted to Firebase!",
    toastFirebaseError: "Firebase submission failed: {message}",
    toastAllInputsCleaned: "All text inputs have been cleaned.",
    // Buttons & Labels
    btnCleanInputs: "Clean All Inputs",
    // Form Fields (Main)
    lblBranchName: "Branch Name:",
    lblRegion: "Region:", // NEW
    // Table Headers
    tblColBranch: "Branch",
    tblColRegion: "Region", // NEW
    // Placeholders (Main)
    phBranchName: "Enter branch name",
    phRegion: "Enter region", // NEW
    // ... (rest of the English translations from your original file)
    toastSpeechUnsupp: "Speech recognition not supported by this browser.",
    toastPartnerLoadErr: "Error loading partner configuration.",
    toastPartnerSaveErr: "Error saving partner configuration.",
    toastCustSaveErr: "Error saving customer data.",
    toastLoadErr: "Error loading previous data. Starting fresh.",
    toastDataCleared: "All customer data cleared.",
    toastHeard: "Heard:",
    toastMicDenied:
      "Microphone access denied. Please allow microphone access in browser settings.",
    toastNoSpeech: "No speech detected. Please try again.",
    toastAudioErr: "Audio capture error. Check microphone connection.",
    toastSpeechErr: "Speech recognition error:",
    toastListening: "Listening...",
    toastGenericErr: "An unexpected error occurred.",
    toastInputCleaned: "Input cleaned.",
    toastImageUploadLimit: "Cannot upload more than {limit} images.",
    toastImageUploadSuccess: "Images uploaded. Total: {count}",
    toastImageUploadErr: "Error uploading image: {name}",
    toastNoImages: "No images uploaded yet.",
    toastReqFieldsWarning: "Please fill all required fields: {fields}",
    toastEntryUpdated: "Entry for {name} updated.",
    toastEntryAdded: "Entry for {name} added.",
    toastEditing: "Editing entry for {name}.",
    DeleteEntry: "Are you sure you want to delete the entry for",
    toastEntryDeleted: "Entry for {name} deleted.",
    toastExcelDownloadOnly: "Data downloaded to Excel.",
    toastExcelErr: "Error exporting to Excel. Ensure XLSX library is loaded.",
    toastCsvSuccess: "Data saved to CSV.",
    toastCsvErr: "Error exporting to CSV.",
    toastPdfSuccess: "Data saved to PDF.",
    toastPdfErr:
      "Error exporting to PDF. Ensure jsPDF & autoTable libraries are loaded.",
    toastDashErr:
      "Error generating dashboard. Ensure Chart.js library is loaded.",
    toastFullscreenErr: "Fullscreen request failed: {message}",
    toastLoginSuccess: "Admin login successful.",
    toastLoginFail: "Invalid username or password.",
    toastAdminSaveValidation:
      "Validation Error: Please check Partner/Product/Tier details (names cannot be empty, premium/tenure must be non-negative numbers).",
    toastAdminSaveSuccess: "Partner configuration saved successfully.",
    toastAdminSaveNoChange: "No changes detected in partner configuration.",
    toastAdminLoadError: "Error loading data for admin panel.",
    toastAdminDeleteConfirm:
      "Are you sure you want to delete this {itemType}: {itemName}?",
    toastAdminItemDeleted: "{itemType} '{itemName}' deleted successfully.",
    toastAdminAddError: "Error adding {itemType}.",
    toastAdminAddSuccess: "{itemType} '{itemName}' added successfully.",
    toastLibNotLoaded: "{libraryName} library not loaded. Feature disabled.",
    btnThemeLight: "Light",
    btnThemeDark: "Dark",
    btnEnterDataEntryMode: "Enter Data Entry Mode",
    lblEnterDataEntryMode: "Focus Mode",
    btnExitDataEntryMode: "Exit Data Entry Mode",
    lblExitDataEntryMode: "View All",
    lblRecordAll: "Record All",
    btnRecordAll: "Record All Fields Sequentially",
    optSelectPartner: "-- Select Partner --",
    optSelectPartnerFirst: "-- Select Partner First --",
    optSelectProduct: "-- Select Product --",
    optSelectProductFirst: "-- Select Product First --",
    optSelectPremium: "-- Select Premium --",
    optSelectPremiumFirst: "-- Select Product First --",
    optSelectRelationship: "-- Select Relationship --",
    btnAddEntry: "Add Entry",
    btnUpdateEntry: "Update Entry",
    btnClearForm: "Clear Form",
    btnSaveExcel: "Download Excel",
    btnSaveCsv: "Download CSV",
    btnSavePdf: "Download PDF",
    btnDashboard: "Dashboard",
    btnClose: "Close",
    btnAdminLogin: "Admin Login",
    btnAdminSaveChanges: "Save Changes",
    btnFullscreen: "Fullscreen",
    btnExitFullscreen: "Exit Fullscreen",
    btnPrev: "Previous",
    btnNext: "Next",
    btnUploadImage: "Upload Images",
    btnToggleImageViewer: "View Images",
    btnClearTable: "Clear Table Data",
    lblPartnerName: "Partner Name:",
    lblProductDetails: "Product Details:",
    lblPremium: "Premium:",
    lblTenure: "Tenure:",
    lblCseName: "CSE Name:",
    lblCustomerName: "Customer Name:",
    lblCustomerID: "Customer ID:",
    lblGender: "Gender:",
    optGenderMale: "Male",
    optGenderFemale: "Female",
    optGenderOther: "Other",
    optGenderUnknown: "Unknown",
    lblMobileNumber: "Mobile Number:",
    lblEnrolmentDate: "Enrolment Date:",
    lblDob: "Date of Birth:",
    lblSavingsAccNum: "Savings A/C No.:",
    lblCsbCode: "CSB Code:",
    lblD2CRoCode: "D2C Code / RO Code:",
    nomineeSectionTitle: "Nominee Details",
    lblNomineeName: "Nominee Name:",
    lblNomineeRelationship: "Nominee Relationship:",
    lblNomineeDob: "Nominee Date of Birth:",
    lblNomineeMobile: "Nominee Mobile Number:",
    lblNomineeGender: "Nominee Gender:",
    lblRemarks: "Remarks:",
    tblColName: "Customer Name",
    tblColCustID: "Customer ID",
    tblColGender: "Gender",
    tblColMobile: "Mobile",
    tblColEnrolDate: "Enrol. Date",
    tblColDOB: "DOB",
    tblColSavingsAcc: "Savings A/C",
    tblColCSB: "CSB Code",
    tblColD2CRoCode: "D2C/RO Code",
    tblColPartnerName: "Partner",
    tblColProductName: "Product",
    tblColPremium: "Premium",
    tblColTenure: "Tenure",
    tblColCseName: "CSE Name",
    tblColNomName: "Nominee Name",
    tblColNomRelation: "Nominee Rel.",
    tblColNomDOB: "Nominee DOB",
    tblColNomMobile: "Nominee Mobile",
    tblColNomGender: "Nominee Gender",
    tblColRemarks: "Remarks",
    tblColActions: "Actions",
    tblNoEntries: "No customer entries yet.",
    mainHeading: "Mswasth Entry Portal",
    tblHeading: "Table",
    lblLanguage: "Language:",
    lblTheme: "Theme:",
    lblAdminLoginTitle: "Admin Login",
    lblUsername: "Username:",
    lblPassword: "Password:",
    lblAdminEditTitle: "Edit Partner & Product Details",
    lblImageUpload: "Image Upload:",
    lblImageCount: "Uploaded Images:",
    ImageViewerTitle: "Image Viewer",
    ImageViewerStatus: "Image {current} of {total}",
    DashboardTitle: "Data Dashboard",
    dashTotalEntries: "Total Entries",
    dashFemaleRatio: "Female Ratio",
    dashBlankEntries: "Entries w/ Blanks",
    DashboardGenderTitle: "Gender Distribution",
    DashboardBranchTitle: "Top 5 Branches by Entries",
    DashboardPartnerTitle: "Entries per Partner",
    phCustomerName: "Enter customer name",
    phCustomerID: "Enter customer ID",
    phMobileNumber: "Enter 10-12 digit number",
    phSavingsAccNum: "Enter savings account number",
    phCsbCode: "Enter CSB code",
    phD2CRoCode: "Enter D2C/RO Code",
    phNomineeName: "Enter Nominee Name",
    phNomineeMobile: "Enter 10-digit mobile number",
    phRemarks: "Enter any remarks or reasons for skipped fields...",
    phAdminUsername: "Enter admin username",
    phAdminPassword: "Enter admin password",
    titleMicButton: "Record",
    titleCleanButton: "Clean",
    titleEditButton: "Edit Entry",
    titleDeleteButton: "Delete Entry",
    titleAdminPartnerName: "Partner Name (Cannot be edited here)",
    titleAdminProductName: "Product Name",
    titleAdminPremium: "Premium Amount (INR)",
    titleAdminTenure: "Tenure (Units: e.g., months/years)",
    titleAdminCseName: "CSE Name",
    titleSaveAdminChanges: "Save All Partner/Product Changes",
    titleCloseModal: "Close",
    titleSpeechLangSelect: "Select Speech Language",
    optRelSpouse: "Spouse",
    optRelSon: "Son",
    optRelDaughter: "Daughter",
    optRelMother: "Mother",
    optRelFather: "Father",
    optRelSister: "Sister",
    optRelBrother: "Brother",
    optRelGrandfather: "Grandfather",
    optRelGrandmother: "Grandmother",
    optRelNephew: "Nephew",
    optRelNiece: "Niece",
    optRelUncle: "Uncle",
    optRelAunty: "Aunty",
    optRelOther: "Other",
    optRelUnknown: "Unknown",
  },
  hi: {
    // Toasts
    toastNothingToSubmit: "सबमिट करने के लिए कोई डेटा नहीं है।",
    toastSubmitSuccess: "{count} प्रविष्टियाँ सफलतापूर्वक फायरबेस में सबमिट की गईं!",
    toastFirebaseError: "फायरबेस सबमिशन विफल: {message}",
    toastAllInputsCleaned: "सभी टेक्स्ट इनपुट साफ़ कर दिए गए हैं।",
    // Buttons & Labels
    btnCleanInputs: "सभी इनपुट साफ़ करें",
    // Form Fields (Main)
    lblBranchName: "शाखा का नाम:",
    lblRegion: "क्षेत्र:", // NEW
    // Table Headers
    tblColBranch: "शाखा",
    tblColRegion: "क्षेत्र", // NEW
    // Placeholders (Main)
    phBranchName: "शाखा का नाम दर्ज करें",
    phRegion: "क्षेत्र दर्ज करें", // NEW
    // ... (rest of the Hindi translations from your original file)
    toastSpeechUnsupp: "यह ब्राउज़र स्पीच रिकग्निशन का समर्थन नहीं करता है।",
    toastPartnerLoadErr: "पार्टनर कॉन्फ़िगरेशन लोड करने में त्रुटि।",
    toastPartnerSaveErr: "पार्टनर कॉन्फ़िगरेशन सहेजने में त्रुटि।",
    toastCustSaveErr: "ग्राहक डेटा सहेजने में त्रुटि।",
    toastLoadErr: "पिछला डेटा लोड करने में त्रुटि। नए सिरे से शुरू कर रहे हैं।",
    toastDataCleared: "सभी ग्राहक डेटा साफ़ कर दिए गए हैं।",
    toastHeard: "सुना गया:",
    toastMicDenied: "माइक्रोफ़ोन एक्सेस अस्वीकृत। कृपया ब्राउज़र सेटिंग्स में माइक्रोफ़ोन एक्सेस की अनुमति दें।",
    toastNoSpeech: "कोई आवाज़ नहीं मिली। कृपया पुनः प्रयास करें।",
    toastAudioErr: "ऑडियो कैप्चर त्रुटि। माइक्रोफ़ोन कनेक्शन जांचें।",
    toastSpeechErr: "स्पीच रिकग्निशन त्रुटि:",
    toastListening: "सुन रहा हूँ...",
    toastGenericErr: "एक अप्रत्याशित त्रुटि हुई।",
    toastInputCleaned: "इनपुट साफ़ किया गया।",
    toastImageUploadLimit: "{limit} से अधिक चित्र अपलोड नहीं कर सकते।",
    toastImageUploadSuccess: "चित्र अपलोड किए गए। कुल: {count}",
    toastImageUploadErr: "चित्र अपलोड करने में त्रुटि: {name}",
    toastNoImages: "अभी तक कोई चित्र अपलोड नहीं किया गया है।",
    toastReqFieldsWarning: "कृपया सभी आवश्यक फ़ील्ड भरें: {fields}",
    toastEntryUpdated: "{name} के लिए प्रविष्टि अपडेट की गई।",
    toastEntryAdded: "{name} के लिए प्रविष्टि जोड़ी गई।",
    toastEditing: "{name} के लिए प्रविष्टि संपादित कर रहे हैं।",
    DeleteEntry: "क्या आप वाकई {name} के लिए प्रविष्टि हटाना चाहते हैं?",
    toastEntryDeleted: "{name} के लिए प्रविष्टि हटा दी गई।",
    toastExcelDownloadOnly: "डेटा एक्सेल में डाउनलोड किया गया।",
    toastExcelErr: "एक्सेल में निर्यात करने में त्रुटि। सुनिश्चित करें कि XLSX लाइब्रेरी लोड हो गई है।",
    toastCsvSuccess: "डेटा CSV में सहेजा गया।",
    toastCsvErr: "CSV में निर्यात करने में त्रुटि।",
    toastPdfSuccess: "डेटा PDF में सहेजा गया।",
    toastPdfErr: "PDF में निर्यात करने में त्रुटि। सुनिश्चित करें कि jsPDF और autoTable लाइब्रेरी लोड हो गई हैं।",
    toastDashErr: "डैशबोर्ड बनाने में त्रुटि। सुनिश्चित करें कि Chart.js लाइब्रेरी लोड हो गई है।",
    toastFullscreenErr: "फ़ुलस्क्रीन अनुरोध विफल: {message}",
    toastLoginSuccess: "एडमिन लॉगिन सफल।",
    toastLoginFail: "अमान्य उपयोगकर्ता नाम या पासवर्ड।",
    toastAdminSaveValidation: "सत्यापन त्रुटि: कृपया पार्टनर/उत्पाद/टियर विवरण जांचें (नाम खाली नहीं हो सकते, प्रीमियम/अवधि गैर-नकारात्मक संख्या होनी चाहिए)।",
    toastAdminSaveSuccess: "पार्टनर कॉन्फ़िगरेशन सफलतापूर्वक सहेजा गया।",
    toastAdminSaveNoChange: "पार्टनर कॉन्फ़िगरेशन में कोई बदलाव नहीं पाया गया।",
    toastAdminLoadError: "एडमिन पैनल के लिए डेटा लोड करने में त्रुटि।",
    toastAdminDeleteConfirm: "क्या आप वाकई इस {itemType}: {itemName} को हटाना चाहते हैं?",
    toastAdminItemDeleted: "{itemType} '{itemName}' सफलतापूर्वक हटा दिया गया।",
    toastAdminAddError: "{itemType} जोड़ने में त्रुटि।",
    toastAdminAddSuccess: "{itemType} '{itemName}' सफलतापूर्वक जोड़ा गया।",
    toastLibNotLoaded: "{libraryName} लाइब्रेरी लोड नहीं हुई है। सुविधा अक्षम है।",
    btnThemeLight: "लाइट",
    btnThemeDark: "डार्क",
    btnEnterDataEntryMode: "डेटा एंट्री मोड दर्ज करें",
    lblEnterDataEntryMode: "फोकस मोड",
    btnExitDataEntryMode: "डेटा एंट्री मोड से बाहर निकलें",
    lblExitDataEntryMode: "सभी देखें",
    lblRecordAll: "सभी रिकॉर्ड करें",
    btnRecordAll: "सभी फ़ील्ड क्रमिक रूप से रिकॉर्ड करें",
    optSelectPartner: "-- पार्टनर चुनें --",
    optSelectPartnerFirst: "-- पहले पार्टनर चुनें --",
    optSelectProduct: "-- उत्पाद चुनें --",
    optSelectProductFirst: "-- पहले उत्पाद चुनें --",
    optSelectPremium: "-- प्रीमियम चुनें --",
    optSelectPremiumFirst: "-- पहले उत्पाद चुनें --",
    optSelectRelationship: "-- संबंध चुनें --",
    btnAddEntry: "प्रविष्टि जोड़ें",
    btnUpdateEntry: "प्रविष्टि अपडेट करें",
    btnClearForm: "फ़ॉर्म साफ़ करें",
    btnSaveExcel: "एक्सेल डाउनलोड करें",
    btnSaveCsv: "CSV के रूप में सहेजें",
    btnSavePdf: "PDF के रूप में सहेजें",
    btnDashboard: "डैशबोर्ड",
    btnClose: "बंद करें",
    btnAdminLogin: "एडमिन लॉगिन",
    btnAdminSaveChanges: "बदलाव सहेजें",
    btnFullscreen: "फ़ुलस्क्रीन",
    btnExitFullscreen: "फ़ुलस्क्रीन से बाहर निकलें",
    btnPrev: "पिछला",
    btnNext: "अगला",
    btnUploadImage: "चित्र अपलोड करें",
    btnToggleImageViewer: "चित्र देखें",
    btnClearTable: "तालिका डेटा साफ़ करें",
    lblPartnerName: "पार्टनर का नाम:",
    lblProductDetails: "उत्पाद विवरण:",
    lblPremium: "प्रीमियम:",
    lblTenure: "अवधि:",
    lblCseName: "CSE नाम:",
    lblCustomerName: "ग्राहक का नाम:",
    lblCustomerID: "ग्राहक आईडी:",
    lblGender: "लिंग:",
    optGenderMale: "पुरुष",
    optGenderFemale: "महिला",
    optGenderOther: "अन्य",
    optGenderUnknown: "अज्ञात",
    lblMobileNumber: "मोबाइल नंबर:",
    lblEnrolmentDate: "नामांकन तिथि:",
    lblDob: "जन्म तिथि:",
    lblSavingsAccNum: "बचत खाता संख्या:",
    lblCsbCode: "सीएसबी कोड:",
    lblD2CRoCode: "D2C कोड / RO कोड:",
    nomineeSectionTitle: "नामित व्यक्ति का विवरण",
    lblNomineeName: "नामित व्यक्ति का नाम:",
    lblNomineeRelationship: "नामित व्यक्ति का संबंध:",
    lblNomineeDob: "नामित व्यक्ति की जन्म तिथि:",
    lblNomineeMobile: "नामित व्यक्ति का मोबाइल नंबर:",
    lblNomineeGender: "नामित व्यक्ति का लिंग:",
    lblRemarks: "टिप्पणियां:",
    tblColName: "ग्राहक का नाम",
    tblColCustID: "ग्राहक आईडी",
    tblColGender: "लिंग",
    tblColMobile: "मोबाइल",
    tblColEnrolDate: "नामांकन तिथि",
    tblColDOB: "जन्म तिथि",
    tblColSavingsAcc: "बचत खाता",
    tblColCSB: "सीएसबी कोड",
    tblColD2CRoCode: "D2C/RO कोड",
    tblColPartnerName: "पार्टनर",
    tblColProductName: "उत्पाद",
    tblColPremium: "प्रीमियम",
    tblColTenure: "अवधि",
    tblColCseName: "CSE नाम",
    tblColNomName: "नामित व्यक्ति का नाम",
    tblColNomRelation: "नामित व्यक्ति का संबंध",
    tblColNomDOB: "नामित व्यक्ति की जन्मतिथि",
    tblColNomMobile: "नामित व्यक्ति का मोबाइल",
    tblColNomGender: "नामित व्यक्ति का लिंग",
    tblColRemarks: "टिप्पणियां",
    tblColActions: "कार्रवाइयाँ",
    tblNoEntries: "अभी तक कोई ग्राहक प्रविष्टि नहीं है।",
    mainHeading: "एमस्वस्थ एंट्री पोर्टल",
    tblHeading: "तालिका",
    lblLanguage: "भाषा:",
    lblTheme: "थीम:",
    lblAdminLoginTitle: "एडमिन लॉगिन",
    lblUsername: "उपयोगकर्ता नाम:",
    lblPassword: "पासवर्ड:",
    lblAdminEditTitle: "पार्टनर और उत्पाद विवरण संपादित करें",
    lblImageUpload: "छवि अपलोड:",
    lblImageCount: "अपलोड की गई छवियाँ:",
    ImageViewerTitle: "छवि दर्शक",
    ImageViewerStatus: "छवि {current} / {total}",
    DashboardTitle: "डेटा डैशबोर्ड",
    dashTotalEntries: "कुल प्रविष्टियाँ",
    dashFemaleRatio: "महिला अनुपात",
    dashBlankEntries: "रिक्तियों वाली प्रविष्टियाँ",
    DashboardGenderTitle: "लिंग वितरण",
    DashboardBranchTitle: "शीर्ष 5 शाखाएँ (प्रविष्टियों के आधार पर)",
    DashboardPartnerTitle: "प्रति पार्टनर प्रविष्टियाँ",
    phCustomerName: "ग्राहक का नाम दर्ज करें",
    phCustomerID: "ग्राहक आईडी दर्ज करें",
    phMobileNumber: "10-12 अंकों का मोबाइल नंबर दर्ज करें",
    phSavingsAccNum: "बचत खाता संख्या दर्ज करें",
    phCsbCode: "सीएसबी कोड दर्ज करें",
    phD2CRoCode: "D2C/RO कोड दर्ज करें",
    phNomineeName: "नामित व्यक्ति का नाम दर्ज करें",
    phNomineeMobile: "10 अंकों का मोबाइल नंबर दर्ज करें",
    phRemarks: "कोई भी टिप्पणी या छोड़े गए फ़ील्ड के कारण दर्ज करें...",
    phAdminUsername: "एडमिन उपयोगकर्ता नाम दर्ज करें",
    phAdminPassword: "एडमिन पासवर्ड दर्ज करें",
    titleMicButton: "रिकॉर्ड करें",
    titleCleanButton: "इनपुट साफ़ करें",
    titleEditButton: "प्रविष्टि संपादित करें",
    titleDeleteButton: "प्रविष्टि हटाएं",
    titleAdminPartnerName: "पार्टनर का नाम (यहाँ संपादित नहीं किया जा सकता)",
    titleAdminProductName: "उत्पाद का नाम",
    titleAdminPremium: "प्रीमियम राशि (INR)",
    titleAdminTenure: "अवधि (इकाई: जैसे महीने/वर्ष)",
    titleAdminCseName: "CSE नाम",
    titleSaveAdminChanges: "सभी पार्टनर/उत्पाद परिवर्तन सहेजें",
    titleCloseModal: "बंद करें",
    titleSpeechLangSelect: "बोलने की भाषा चुनें",
    optRelSpouse: "पति/पत्नी",
    optRelSon: "बेटा",
    optRelDaughter: "बेटी",
    optRelMother: "माँ",
    optRelFather: "पिता",
    optRelSister: "बहन",
    optRelBrother: "भाई",
    optRelGrandfather: "दादा/नाना",
    optRelGrandmother: "दादी/नानी",
    optRelNephew: "भतीजा/भांजा",
    optRelNiece: "भतीजी/भांजी",
    optRelUncle: "चाचा/मामा/फूफा/मौसा",
    optRelAunty: "चाची/मामी/बुआ/मौसी",
    optRelOther: "अन्य",
    optRelUnknown: "अज्ञात",
  },
};
let currentLang = "en";

// --- DOM Element Selectors (Cached for performance) ---
let domElements = {};
function cacheDomElements() {
  domElements = {
    form: document.getElementById("customerForm"),
    tableBody: document.getElementById("dataTableBody"),
    partnerNameSelect: document.getElementById("partnerNameSelect"),
    productSelect: document.getElementById("productSelect"),
    premiumSelect: document.getElementById("premiumSelect"),
    nomineeRelationshipSelect: document.getElementById("nomineeRelationship"),
    remarksTextarea: document.getElementById("remarks"),
    d2cRoCodeInput: document.getElementById("d2cRoCode"),
    regionInput: document.getElementById("region"), // NEW
    selectedTenureInput: document.getElementById("selectedTenure"),
    selectedCseNameInput: document.getElementById("selectedCseName"),
    nomineeNameInput: document.getElementById("nomineeName"),
    nomineeMobileInput: document.getElementById("nomineeMobileNumber"),
    editingIdInput: document.getElementById("editingId"),
    addEntryButton: document.getElementById("addEntryButton"),
    clearFormButton: document.getElementById("clearFormButton"),
    clearInputsButton: document.getElementById("clearInputsButton"), // NEW
    saveExcelButton: document.getElementById("saveExcelButton"),
    saveCsvButton: document.getElementById("saveCsvButton"),
    savePdfButton: document.getElementById("savePdfButton"),
    languageSwitcher: document.getElementById("languageSwitcher"),
    themeToggleButton: document.getElementById("themeToggleButton"),
    languageSelect: document.getElementById("speechLangSelect"),
    loadingIndicator: document.getElementById("loadingIndicator"),
    dashboardButton: document.getElementById("dashboardButton"),
    dashboardModal: document.getElementById("dashboardModal"),
    closeDashboardButton: document.getElementById("closeDashboardButton"),
    dashboardCarouselSlides: document.getElementById("dashboardCarouselSlides"),
    prevChartButton: document.getElementById("prevChartButton"),
    nextChartButton: document.getElementById("nextChartButton"),
    adminButton: document.getElementById("adminButton"),
    adminLoginModal: document.getElementById("adminLoginModal"),
    closeAdminLoginButton: document.getElementById("closeAdminLoginButton"),
    adminLoginForm: document.getElementById("adminLoginForm"),
    adminUsernameInput: document.getElementById("adminUsername"),
    adminPasswordInput: document.getElementById("adminPassword"),
    adminEditModal: document.getElementById("adminEditModal"),
    closeAdminEditButton: document.getElementById("closeAdminEditButton"),
    adminEditForm: document.getElementById("adminEditForm"),
    adminEditGrid: document.getElementById("adminEditGrid"),
    fullscreenButton: document.getElementById("fullscreenButton"),
    tableContainer: document.getElementById("tableContainer"),
    uploadImageButton: document.getElementById("uploadImageButton"),
    imageUploadInput: document.getElementById("imageUploadInput"),
    imageCountSpan: document.getElementById("imageCount"),
    toggleImageViewerButton: document.getElementById("toggleImageViewerButton"),
    imageViewerSection: document.getElementById("imageViewerSection"),
    submitToFirebaseButton: document.getElementById("submitToFirebaseButton"),
    currentImageView: document.getElementById("currentImageView"),
    imageViewerStatus: document.getElementById("imageViewerStatus"),
    prevImageButton: document.getElementById("prevImageButton"),
    nextImageButton: document.getElementById("nextImageButton"),
    toggleDataEntryModeButton: document.getElementById(
      "toggleDataEntryModeButton"
    ),
    dataTableHead: document.querySelector("#dataTable thead"),
    recordAllButton: document.getElementById("recordAllButton"),
    clearTableButton: document.getElementById("clearTableButton"),
    resetApplicationDataButton: document.getElementById("resetApplicationDataButton")
  };
}

// --- Initialization ---
document.addEventListener("DOMContentLoaded", () => {
  cacheDomElements();
  showLoader();
  initializeTheme();
  initializeDatePickers(); // NEW
  loadPartnerData();
  loadDataFromLocalStorage();
  initializeLanguage();

  if (!window.SpeechRecognition) {
    showToast(translate("toastSpeechUnsupp"), "error", 5000);
    disableSpeechFunctionality();
  } else {
    initializeSpeechRecognition();
  }

  setupEventListeners();
  updateDataEntryModeButton();
  applyRequiredAttributes();
  renderTable();
  hideLoader();
});

// --- Apply Required Attributes ---
function applyRequiredAttributes() {
    if (!domElements.form) return;
  
    domElements.form.querySelectorAll("input[required], select[required]").forEach((el) => el.removeAttribute("required"));
    domElements.form.querySelectorAll('input[type="radio"][required]').forEach((el) => el.removeAttribute("required"));
  
    requiredFormElementIds.forEach((id) => {
      const element = document.getElementById(id);
      if (element) element.setAttribute("required", "");
    });
  
    requiredRadioGroups.forEach((groupName) => {
      const firstRadio = domElements.form.querySelector(`input[name="${groupName}"]`);
      if (firstRadio) firstRadio.setAttribute("required", "");
    });
}

// --- NEW: Initialize Flatpickr Date Pickers ---
function initializeDatePickers() {
  if (typeof flatpickr === 'undefined') {
    console.error("Flatpickr library not loaded.");
    return;
  }
  const datePickers = document.querySelectorAll('.datepicker-input');
  flatpickrInstances = Array.from(datePickers).map(picker => {
    return flatpickr(picker, {
      dateFormat: "d-m-Y",
      allowInput: true, // Allows manual typing
    });
  });
}

// --- Internationalization (i18n) ---
function translate(key, replacements = {}) {
  const langTranslations = translations[currentLang] || translations.en;
  let translation = langTranslations[key];

  if (translation === undefined) {
    translation = translations.en[key];
  }
  if (translation === undefined) {
    return key;
  }

  for (const placeholder in replacements) {
    const regex = new RegExp(`\\{${placeholder}\\}`, "g");
    translation = translation.replace(regex, replacements[placeholder]);
  }
  return translation;
}

function setLanguage(lang) {
    if (!translations[lang]) {
      lang = "en";
    }
    currentLang = lang;
    document.documentElement.lang = lang;
    localStorage.setItem(LS_UI_LANG_KEY, lang);
  
    document.querySelectorAll("[data-translate-key]").forEach((element) => {
      const key = element.dataset.translateKey;
      let text = translate(key);
      const isPlaceholderOption = element.tagName === "OPTION" && element.value === "";
  
      if (element.placeholder !== undefined && element.dataset.translatePlaceholderKey !== "false") {
        const placeholderKey = element.dataset.translatePlaceholderKey || `ph${key.substring(3)}`;
        let actualPlaceholderKey = placeholderKey;
        if (key.startsWith("lbl") && !translations[currentLang]?.[placeholderKey] && !translations.en?.[placeholderKey]) {
          actualPlaceholderKey = key;
        }
        element.placeholder = translate(actualPlaceholderKey, {}, key);
      }
  
      if (element.title !== undefined && element.dataset.translateTitleKey !== "false") {
        const titleKey = element.dataset.translateTitleKey || key;
        let titleReplacements = {};
        if (titleKey === "titleMicButton" || titleKey === "titleCleanButton") {
          const targetId = element.dataset.target;
          const label = domElements.form?.querySelector(`label[for='${targetId}']`);
          let fieldName = targetId || "field";
          if (label) {
            const labelClone = label.cloneNode(true);
            labelClone.querySelectorAll(".required-star, .sr-only").forEach((s) => s.remove());
            fieldName = labelClone.textContent.replace(/[:*]/g, "").trim();
          } else {
            const labelKey = `lbl${targetId.charAt(0).toUpperCase() + targetId.slice(1)}`;
            const translatedLabel = translate(labelKey);
            if (translatedLabel !== labelKey) fieldName = translatedLabel.replace(/[:*]/g, "").trim();
          }
          titleReplacements = { fieldName: fieldName || targetId };
        }
        element.title = translate(titleKey, titleReplacements);
      }
  
      let targetElementForText = element;
      const specificTextSpan = element.querySelector("span:not(.sr-only):not(.sort-icon):not(.required-star)");
      if (specificTextSpan && !element.matches(".stat-card, .gender-group label, .control-button")) {
        targetElementForText = specificTextSpan;
      }
  
      if (!["INPUT", "SELECT", "TEXTAREA", "IMG"].includes(element.tagName)) {
        if (isPlaceholderOption) {
          element.textContent = text;
        } else if (targetElementForText === element && element.children.length > 0 && element.firstChild?.nodeType === Node.TEXT_NODE) {
          element.firstChild.textContent = text + " ";
        } else if (targetElementForText.childNodes.length === 1 && targetElementForText.firstChild?.nodeType === Node.TEXT_NODE) {
          targetElementForText.textContent = text;
        } else if (targetElementForText !== element) {
          targetElementForText.textContent = text;
        } else if (targetElementForText.childNodes.length === 0) {
          targetElementForText.textContent = text;
        }
      }
    });
  
    applyTheme(localStorage.getItem(LS_THEME_KEY) || "light-theme");
    updateFullscreenButtonText();
    updateDataEntryModeButton();
    if (domElements.languageSwitcher) domElements.languageSwitcher.value = lang;
    if (domElements.languageSelect) {
      let speechLang = lang === "hi" ? "hi-IN" : "en-US";
      if (Array.from(domElements.languageSelect.options).some((opt) => opt.value === speechLang)) {
        domElements.languageSelect.value = speechLang;
      } else {
        domElements.languageSelect.value = "en-US";
      }
    }
    populatePartnerNameDropdown();
    applyRequiredAttributes();
    renderTable();
}

function initializeLanguage() {
  const savedLang =
    localStorage.getItem(LS_UI_LANG_KEY) ||
    (navigator.language || "en").split("-")[0] ||
    "en";
  setLanguage(savedLang);
}

// --- Theme Management ---
function initializeTheme() {
  const theme = localStorage.getItem(LS_THEME_KEY) || "light-theme";
  applyTheme(theme);
}
// MODIFIED: applyTheme now also handles the Flatpickr theme
function applyTheme(theme) {
  document.body.classList.remove("light-theme", "dark-theme");
  document.body.classList.add(theme);
  localStorage.setItem(LS_THEME_KEY, theme);

  // Switch Flatpickr theme
  const flatpickrTheme = theme === "dark-theme" ? "dark" : "light";
  document.querySelectorAll('.flatpickr-calendar').forEach(cal => {
    cal.classList.remove('dark', 'light');
    cal.classList.add(flatpickrTheme);
  });

  if (!domElements.themeToggleButton) return;
  const icon = domElements.themeToggleButton.querySelector("i");
  const textSpan = domElements.themeToggleButton.querySelector("span");
  if (icon && textSpan) {
    if (theme === "dark-theme") {
      icon.className = "fas fa-moon";
      textSpan.textContent = translate("btnThemeDark");
    } else {
      icon.className = "fas fa-sun";
      textSpan.textContent = translate("btnThemeLight");
    }
  }
}
function toggleTheme() {
  const currentTheme = document.body.classList.contains("dark-theme")
    ? "dark-theme"
    : "light-theme";
  const newTheme = currentTheme === "dark-theme" ? "light-theme" : "dark-theme";
  applyTheme(newTheme);
  showToast(`Theme changed to ${newTheme.replace("-theme", "")}`, "info");
}

// --- Data Entry Mode ---
function toggleDataEntryMode() {
  document.body.classList.toggle("data-entry-mode");
  updateDataEntryModeButton();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function updateDataEntryModeButton() {
    if (!domElements.toggleDataEntryModeButton) return;
    const isInEntryMode = document.body.classList.contains("data-entry-mode");
    const button = domElements.toggleDataEntryModeButton;
    const icon = button.querySelector("i");
    const span = button.querySelector("span");
  
    if (isInEntryMode) {
      button.title = translate("btnExitDataEntryMode");
      if (span) span.textContent = translate("lblExitDataEntryMode");
      if (icon) icon.className = "fas fa-table";
    } else {
      button.title = translate("btnEnterDataEntryMode");
      if (span) span.textContent = translate("lblEnterDataEntryMode");
      if (icon) icon.className = "fas fa-columns";
    }
}

// --- Partner/Product/Premium Data Management ---
function loadPartnerData() {
    const storedData = localStorage.getItem(LS_PARTNER_KEY);
    if (storedData) {
      try {
        const parsed = JSON.parse(storedData);
        if (typeof parsed === "object" && parsed !== null && Object.keys(parsed).length > 0) {
          partnerData = parsed;
        } else {
          savePartnerData();
        }
      } catch (e) {
        showToast(translate("toastPartnerLoadErr"), "error");
      }
    } else {
      savePartnerData();
    }
}

function savePartnerData() {
  try {
    localStorage.setItem(LS_PARTNER_KEY, JSON.stringify(partnerData));
    console.log("Partner data saved to localStorage.");
  } catch (e) {
    console.error("Error saving partner data to localStorage:", e);
    showToast(translate("toastPartnerSaveErr"), "error");
  }
}

function populatePartnerNameDropdown() {
  if (!domElements.partnerNameSelect) return;
  const select = domElements.partnerNameSelect;
  const currentVal = select.value;
  select.innerHTML = `<option value="">${translate(
    "optSelectPartner"
  )}</option>`;

  const partnerNames = Object.keys(partnerData).sort();
  partnerNames.forEach((name) => {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name;
    select.appendChild(option);
  });

  if (partnerNames.includes(currentVal)) {
    select.value = currentVal;
  } else {
    select.value = "";
  }

  if (select.value === "") {
    clearProductDropdown();
  } else if (select.value === currentVal) {
    const currentProduct = domElements.productSelect?.value;
    if (!currentProduct || !partnerData[select.value]?.[currentProduct]) {
      clearProductDropdown();
    } else {
      const currentPremium = domElements.premiumSelect?.value;
      const tiers = partnerData[select.value]?.[currentProduct];
      const premiumExists = tiers?.some(
        (tier) => tier.Premium.toString() === currentPremium
      );
      if (!currentPremium || !premiumExists) {
        clearPremiumDropdown();
      }
    }
  }
  applyRequiredAttributes();
}

function populateProductDropdown(selectedPartnerName) {
  if (
    !domElements.productSelect ||
    !selectedPartnerName ||
    !partnerData[selectedPartnerName]
  ) {
    clearProductDropdown();
    return;
  }
  const select = domElements.productSelect;
  const products = partnerData[selectedPartnerName];
  const currentVal = select.value;
  const productNames = Object.keys(products).sort();

  select.innerHTML = `<option value="">${translate(
    "optSelectProduct"
  )}</option>`;
  select.disabled = false;

  productNames.forEach((name) => {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name;
    select.appendChild(option);
  });

  if (productNames.includes(currentVal)) {
    select.value = currentVal;
  } else {
    select.value = "";
  }

  if (select.value === currentVal && currentVal !== "") {
    select.dispatchEvent(new Event("change"));
  } else {
    clearPremiumDropdown();
  }
  applyRequiredAttributes();
}

function populatePremiumDropdown(selectedPartnerName, selectedProductName) {
  if (
    !domElements.premiumSelect ||
    !selectedPartnerName ||
    !selectedProductName ||
    !partnerData[selectedPartnerName] ||
    !partnerData[selectedPartnerName][selectedProductName]
  ) {
    clearPremiumDropdown();
    return;
  }

  const select = domElements.premiumSelect;
  const tiers = partnerData[selectedPartnerName][selectedProductName];
  const currentVal = select.value;
  select.innerHTML = `<option value="">${translate(
    "optSelectPremium"
  )}</option>`;
  select.disabled = false;

  const sortedTiers = [...tiers].sort((a, b) => a.Premium - b.Premium);
  const validPremiums = sortedTiers.map((t) => t.Premium.toString());

  sortedTiers.forEach((tier) => {
    const option = document.createElement("option");
    option.value = tier.Premium.toString();
    option.textContent = formatCurrency(tier.Premium);
    select.appendChild(option);
  });

  if (validPremiums.includes(currentVal)) {
    select.value = currentVal;
  } else {
    select.value = "";
  }

  if (select.value === currentVal && currentVal !== "") {
    select.dispatchEvent(new Event("change"));
  } else {
    updateAutoSelectedFields("", "", "");
  }
  applyRequiredAttributes();
}

function updateAutoSelectedFields(
  partnerName,
  productName,
  selectedPremiumValue
) {
  if (!domElements.selectedTenureInput || !domElements.selectedCseNameInput)
    return;
  let tenure = "";
  let cseName = "";
  if (partnerName && productName && selectedPremiumValue) {
    const premiumNum = parseFloat(selectedPremiumValue);
    if (!isNaN(premiumNum)) {
      const tiers = partnerData[partnerName]?.[productName];
      const selectedTier = tiers?.find((tier) => tier.Premium === premiumNum);
      if (selectedTier) {
        tenure = selectedTier.Tenure;
        cseName = selectedTier["CSE Name"];
      }
    }
  }
  domElements.selectedTenureInput.value = tenure;
  domElements.selectedCseNameInput.value = cseName;
}

function clearProductDropdown() {
  if (!domElements.productSelect) return;
  domElements.productSelect.innerHTML = `<option value="">${translate(
    "optSelectPartnerFirst"
  )}</option>`;
  domElements.productSelect.disabled = true;
  domElements.productSelect.value = "";
  clearPremiumDropdown();
  applyRequiredAttributes();
}

function clearPremiumDropdown() {
  if (!domElements.premiumSelect) return;
  domElements.premiumSelect.innerHTML = `<option value="">${translate(
    "optSelectProductFirst"
  )}</option>`;
  domElements.premiumSelect.disabled = true;
  domElements.premiumSelect.value = "";
  updateAutoSelectedFields("", "", "");
  applyRequiredAttributes();
}

function formatCurrency(amount) {
  if (amount == null || isNaN(amount)) return "";
  return Number(amount).toLocaleString("en-IN", {
    style: "currency",
    currency: "INR",
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  });
}

// --- Customer Data Persistence ---
function saveDataToLocalStorage() {
  try {
    localStorage.setItem(LS_DATA_KEY, JSON.stringify(collectedData));
  } catch (e) {
    console.error("Error saving customer data to localStorage:", e);
    showToast(translate("toastCustSaveErr"), "error");
  }
}

function loadDataFromLocalStorage() {
  const storedData = localStorage.getItem(LS_DATA_KEY);
  if (storedData) {
    try {
      const parsedData = JSON.parse(storedData);
      if (Array.isArray(parsedData)) {
        collectedData = parsedData;
      } else {
        collectedData = [];
      }
    } catch (e) {
      console.error("Error parsing customer data from localStorage:", e);
      collectedData = [];
      showToast(translate("toastLoadErr"), "warning");
    }
  } else {
    collectedData = [];
  }
}

function clearLocalStorageAndCustomerData() {
  localStorage.removeItem(LS_DATA_KEY);
  collectedData = [];
  renderTable();
  showToast(translate("toastDataCleared"), "success");
}

// --- UI Feedback ---
function showToast(message, type = "info", duration = 3000) {
  if (typeof Toastify === "undefined") {
    console.warn("Toastify library not loaded. Message:", message);
    alert(`${type.toUpperCase()}: ${message}`);
    return;
  }
  let background;
  switch (type) {
    case "success":
      background = "linear-gradient(to right, #00b09b, #96c93d)";
      break;
    case "error":
      background = "linear-gradient(to right, #ff5f6d, #ffc371)";
      break;
    case "warning":
      background = "linear-gradient(to right, #f7b733, #fc4a1a)";
      break;
    default:
      background = "linear-gradient(to right, #007bff, #0056b3)";
  }
  Toastify({
    text: message,
    duration: duration,
    close: true,
    gravity: "top",
    position: "right",
    stopOnFocus: true,
    style: { background: background, borderRadius: "5px", zIndex: 10000 },
  }).showToast();
}
function showLoader() {
  if (domElements.loadingIndicator)
    domElements.loadingIndicator.style.display = "flex";
}
function hideLoader() {
  if (domElements.loadingIndicator)
    domElements.loadingIndicator.style.display = "none";
}

// --- Speech Recognition ---
function initializeSpeechRecognition() {
  try {
    recognition = new window.SpeechRecognition();
    recognition.continuous = false;
    recognition.interimResults = false;
    recognition.onstart = handleRecognitionStart;
    recognition.onresult = handleRecognitionResult;
    recognition.onspeechend = () => recognition.stop();
    recognition.onend = handleRecognitionEnd;
    recognition.onerror = handleRecognitionError;
  } catch (e) {
    showToast(translate("toastSpeechUnsupp"), "error", 5000);
    disableSpeechFunctionality();
  }
}

function disableSpeechFunctionality() {
  document.querySelectorAll(".mic-button").forEach((btn) => (btn.disabled = true));
  if (domElements.languageSelect) domElements.languageSelect.disabled = true;
  if (domElements.recordAllButton) domElements.recordAllButton.disabled = true;
}

function handleRecognitionStart() {
  const btn = document.querySelector(`.mic-button[data-target="${currentTargetInput?.id}"]`);
  if (btn) btn.classList.add("listening");
}

function handleRecognitionResult(event) {
  const transcript = event.results[0][0].transcript.trim();
  showToast(`${translate("toastHeard")} "${transcript}"`, "info");
  if (currentTargetInput) {
    // If it's a mobile field, process it to keep only numbers
    if (currentTargetInput.id === 'mobileNumber' || currentTargetInput.id === 'nomineeMobileNumber') {
        currentTargetInput.value = transcript.replace(/\D/g, '');
    } else {
        currentTargetInput.value = transcript;
    }
    currentTargetInput.dispatchEvent(new Event("input", { bubbles: true }));
    currentTargetInput.dispatchEvent(new Event("change", { bubbles: true }));
  }
  if (isRecordingAllActive) {
    removeFieldHighlight(currentTargetInput?.id);
    currentRecordAllIndex++;
    setTimeout(processNextRecordAllField, 100);
  }
}

function handleRecognitionEnd() {
  if (!isRecordingAllActive) {
    enableAllMicButtons();
    removeFieldHighlight(currentTargetInput?.id);
    currentTargetInput = null;
    clearRecordAllStatus();
  }
}


function handleRecognitionError(event) {
  let toastKey = "toastSpeechErr";
  let errorDetail = event.error;
  const showErrorToast = !(isRecordingAllActive && event.error === "no-speech");

  switch (event.error) {
    case "not-allowed":
    case "service-not-allowed":
      toastKey = "toastMicDenied";
      break;
    case "no-speech":
      toastKey = "toastNoSpeech";
      break;
    case "audio-capture":
      toastKey = "toastAudioErr";
      break;
    case "aborted":
      if (isRecordingAllActive) stopRecordAllSequence("Recording cancelled.");
      else handleRecognitionEnd();
      return;
    case "network":
      errorDetail = "Network error";
      break;
    case "bad-grammar":
      errorDetail = "Grammar error";
      break;
    case "language-not-supported":
      errorDetail = "Language not supported";
      break;
  }

  if (showErrorToast) {
    showToast(`${translate(toastKey)} (${errorDetail})`, "error");
  }

  if (isRecordingAllActive) {
    removeFieldHighlight(currentTargetInput?.id);
    currentRecordAllIndex++;
    setTimeout(processNextRecordAllField, 100);
  } else {
    handleRecognitionEnd();
  }
}

function handleMicButtonClick(event) {
  if (isRecordingAllActive) {
    showToast(
      "Please wait for 'Record All' to finish or cancel it.",
      "warning"
    );
    return;
  }
  const targetId = event.currentTarget.dataset.target;
  currentTargetInput = document.getElementById(targetId);
  if (!currentTargetInput) {
    console.error(`Target input '${targetId}' not found.`);
    return;
  }
  if (!recognition) {
    showToast(translate("toastSpeechUnsupp"), "error");
    return;
  }

  try {
    recognition.stop();
  } catch (e) {
    /* Ignore */
  }
  enableAllMicButtons();

  recognition.lang = domElements.languageSelect?.value || "en-US";
  try {
    disableAllMicButtons();
    const currentMicButton = event.currentTarget;
    currentMicButton.classList.add("listening");
    currentMicButton.disabled = true;
    recognition.start();
    showToast(translate("toastListening"), "info", 1500);
  } catch (error) {
    console.error("Error starting speech recognition:", error);
    let startErrorKey = "toastGenericErr";
    if (error.name === "InvalidStateError")
      startErrorKey = "Speech recognition is already active.";
    else startErrorKey += " Could not start listening.";
    showToast(translate(startErrorKey), "error");
    handleRecognitionEnd();
  }
}

// --- Input Cleaning ---
function handleCleanButtonClick(event) {
    const targetId = event.currentTarget.dataset.target;
    const input = document.getElementById(targetId);
    if (input) {
      handleCleanInput(input);
      showToast(translate("toastInputCleaned"), "success", 1500);
    }
}

// NEW: Clean all text inputs at once
function handleCleanAllInputs() {
    textInputIds.forEach(id => {
        const inputElement = document.getElementById(id);
        if(inputElement) {
            handleCleanInput(inputElement);
        }
    });
    showToast(translate('toastAllInputsCleaned'), 'success');
}

// MODIFIED: handleCleanInput is now smarter for mobile numbers
function handleCleanInput(inputElement) {
  if (!inputElement || typeof inputElement.value !== "string") return;
  
  let cleanedValue;
  if (inputElement.id === 'mobileNumber' || inputElement.id === 'nomineeMobileNumber') {
    // For mobile numbers, strip all non-digit characters
    cleanedValue = inputElement.value.replace(/\D/g, '');
  } else {
    // For other inputs, original logic
    cleanedValue = inputElement.value
      .trim()
      .toLowerCase()
      .replace(/\s+/g, "");
  }

  inputElement.value = cleanedValue;
  inputElement.dispatchEvent(new Event("input", { bubbles: true }));
  inputElement.dispatchEvent(new Event("change", { bubbles: true }));
}


// --- Image Handling ---
function handleImageUploadTrigger() {
  if (domElements.imageUploadInput) domElements.imageUploadInput.click();
  else console.error("Image upload input element not found.");
}
function handleImageFiles(event) {
  const files = event.target.files;
  if (!files || files.length === 0 || !domElements.imageUploadInput) return;
  const currentCount = uploadedImages.length;
  const remainingSlots = MAX_IMAGE_UPLOADS - currentCount;
  if (remainingSlots <= 0) {
    showToast(
      translate("toastImageUploadLimit", { limit: MAX_IMAGE_UPLOADS }),
      "warning"
    );
    domElements.imageUploadInput.value = "";
    return;
  }
  const filesToProcess = Array.from(files).slice(0, remainingSlots);
  if (files.length > remainingSlots) {
    showToast(
      translate("toastImageUploadLimit", { limit: MAX_IMAGE_UPLOADS }),
      "warning"
    );
  }
  let processedCount = 0;
  let successfulUploads = 0;
  const totalToProcess = filesToProcess.length;
  if (totalToProcess > 0) showLoader();
  filesToProcess.forEach((file) => {
    if (!file.type.startsWith("image/")) {
      processedCount++;
      if (processedCount === totalToProcess) hideLoader();
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      uploadedImages.push({ name: file.name, dataUrl: e.target.result });
      processedCount++;
      successfulUploads++;
      if (processedCount === totalToProcess) {
        hideLoader();
        updateImageViewerState();
        showToast(
          translate("toastImageUploadSuccess", { count: successfulUploads }),
          "success"
        );
        if (
          domElements.imageViewerSection?.style.display === "none" &&
          uploadedImages.length > 0
        ) {
          toggleImageViewer();
        }
      }
    };
    reader.onerror = () => {
      showToast(translate("toastImageUploadErr", { name: file.name }), "error");
      processedCount++;
      if (processedCount === totalToProcess) hideLoader();
    };
    reader.readAsDataURL(file);
  });
  domElements.imageUploadInput.value = "";
}
function toggleImageViewer() {
  if (!domElements.imageViewerSection || !domElements.toggleImageViewerButton)
    return;
  if (uploadedImages.length === 0) {
    showToast(translate("toastNoImages"), "info");
    domElements.imageViewerSection.style.display = "none";
    domElements.toggleImageViewerButton.disabled = true;
    return;
  }
  const isHidden = domElements.imageViewerSection.style.display === "none";
  domElements.imageViewerSection.style.display = isHidden ? "flex" : "none";
  if (isHidden) {
    currentImageIndex = 0;
    displayCurrentImage();
  }
}
function displayCurrentImage() {
  if (
    !domElements.imageViewerSection ||
    !domElements.currentImageView ||
    !domElements.imageViewerStatus ||
    !domElements.prevImageButton ||
    !domElements.nextImageButton
  )
    return;
  const totalImages = uploadedImages.length;
  if (
    totalImages === 0 ||
    currentImageIndex < 0 ||
    currentImageIndex >= totalImages ||
    !uploadedImages[currentImageIndex]
  ) {
    if (domElements.imageViewerSection)
      domElements.imageViewerSection.style.display = "none";
    if (domElements.currentImageView) {
      domElements.currentImageView.src = "";
      domElements.currentImageView.alt = "";
    }
    if (domElements.imageViewerStatus)
      domElements.imageViewerStatus.textContent = translate(
        "ImageViewerStatus",
        { current: 0, total: 0 }
      );
    if (domElements.prevImageButton)
      domElements.prevImageButton.disabled = true;
    if (domElements.nextImageButton)
      domElements.nextImageButton.disabled = true;
    return;
  }
  if (domElements.imageViewerSection.style.display === "none") {
    domElements.imageViewerSection.style.display = "flex";
  }
  domElements.currentImageView.src = uploadedImages[currentImageIndex].dataUrl;
  domElements.currentImageView.alt =
    uploadedImages[currentImageIndex].name || `Image ${currentImageIndex + 1}`;
  domElements.imageViewerStatus.textContent = translate("ImageViewerStatus", {
    current: currentImageIndex + 1,
    total: totalImages,
  });
  domElements.prevImageButton.disabled = currentImageIndex === 0;
  domElements.nextImageButton.disabled = currentImageIndex === totalImages - 1;
}
function showNextImage() {
  if (currentImageIndex < uploadedImages.length - 1) {
    currentImageIndex++;
    displayCurrentImage();
  }
}
function showPrevImage() {
  if (currentImageIndex > 0) {
    currentImageIndex--;
    displayCurrentImage();
  }
}
function updateImageViewerState() {
  if (!domElements.imageCountSpan || !domElements.toggleImageViewerButton)
    return;
  const count = uploadedImages.length;
  domElements.imageCountSpan.textContent = count;
  domElements.toggleImageViewerButton.disabled = count === 0;
  if (count === 0) {
    if (domElements.imageViewerSection)
      domElements.imageViewerSection.style.display = "none";
    displayCurrentImage();
  } else {
    if (domElements.imageViewerSection?.style.display !== "none") {
      if (currentImageIndex >= count) {
        currentImageIndex = Math.max(0, count - 1);
      }
      displayCurrentImage();
    }
  }
}


// --- Core Customer Data Handling & Validation ---

// NEW: Advanced form validation with live highlighting
function validateAndHighlightForm() {
    if (!domElements.form) return { isValid: false, missingFields: ["Form not found"] };

    let isValid = true;
    const missingFields = [];

    // Check all required inputs, selects, textareas
    domElements.form.querySelectorAll('[required]').forEach(el => {
        const isInvalid = !el.value.trim();
        el.classList.toggle('input-invalid', isInvalid);
        if (isInvalid) {
            isValid = false;
            const label = domElements.form.querySelector(`label[for='${el.id}']`);
            let fieldName = el.id;
            if (label) {
                const labelClone = label.cloneNode(true);
                labelClone.querySelectorAll(".required-star, .sr-only").forEach(s => s.remove());
                fieldName = labelClone.textContent.replace(/[:*]/g, "").trim();
            }
            missingFields.push(fieldName);
        }
    });

    // Check required radio button groups
    requiredRadioGroups.forEach(groupName => {
        const group = domElements.form.elements[groupName];
        if (group && !group.value) {
            isValid = false;
            const groupContainer = domElements.form.querySelector(`input[name='${groupName}']`).closest('.form-field');
            if (groupContainer) {
                const label = groupContainer.querySelector('label');
                missingFields.push(label ? label.textContent.replace(":", "") : groupName);
                groupContainer.querySelector('.gender-group').classList.add('input-invalid');
            }
        } else if (group && group.value) {
            const groupContainer = domElements.form.querySelector(`input[name='${groupName}']`).closest('.form-field');
            if (groupContainer) {
                groupContainer.querySelector('.gender-group').classList.remove('input-invalid');
            }
        }
    });

    return { isValid, missingFields: [...new Set(missingFields)] };
}


function handleAddOrUpdateEntry() {
  if (!domElements.form) return;
  
  const validationResult = validateAndHighlightForm();
  if (!validationResult.isValid) {
    showToast(translate("toastReqFieldsWarning", { fields: validationResult.missingFields.join(", ") }), "warning", 5000);
    const firstInvalidElement = domElements.form.querySelector('.input-invalid');
    if (firstInvalidElement) {
        // Find the actual input if the container was marked as invalid (e.g., gender group)
        (firstInvalidElement.querySelector('input, select, textarea') || firstInvalidElement).focus();
    }
    return;
  }

  const entryData = {};
  customerFieldKeys.forEach((key) => {
    let element;
    let value = null;
    if (key === "gender" || key === "nomineeGender") {
      const checkedRadio = domElements.form.querySelector(`input[name="${key}"]:checked`);
      value = checkedRadio ? checkedRadio.value : "";
    } else if (key === "partnerName") {
      value = domElements.partnerNameSelect ? domElements.partnerNameSelect.value : "";
    } else if (key === "productName") {
      value = domElements.productSelect ? domElements.productSelect.value : "";
    } else if (key === "premium") {
      value = domElements.premiumSelect ? parseFloat(domElements.premiumSelect.value) : null;
      if (isNaN(value)) value = null;
    } else if (key === "nomineeRelationship") {
      value = domElements.nomineeRelationshipSelect ? domElements.nomineeRelationshipSelect.value : "";
    } else if (key === "remarks") {
        value = domElements.remarksTextarea ? domElements.remarksTextarea.value.trim() : "";
    }
    else {
      element = document.getElementById(key);
      if (element) value = element.value.trim();
    }
    entryData[key] = value;
  });

  const editingId = domElements.editingIdInput.value;
  const name = entryData.customerName || "N/A";

  if (editingId) {
    const index = collectedData.findIndex((e) => e.id == editingId);
    if (index > -1) {
      collectedData[index] = { ...collectedData[index], ...entryData };
      showToast(translate("toastEntryUpdated", { name: name }), "success");
    } else {
      entryData.id = Date.now();
      collectedData.push(entryData);
      showToast(translate("toastEntryAdded", { name: name }), "success");
    }
    domElements.editingIdInput.value = "";
    domElements.addEntryButton.querySelector("span").textContent = translate("btnAddEntry");
  } else {
    entryData.id = Date.now();
    collectedData.push(entryData);
    showToast(translate("toastEntryAdded", { name: name }), "success");
  }

  saveDataToLocalStorage();
  renderTable();
  clearForm();
}


function handleEditEntry(id) {
    const entry = collectedData.find((e) => e.id == id);
    if (!entry) {
      showToast("Could not find the entry to edit.", "error");
      return;
    }
    if (!domElements.form) return;
    clearFormValidationStyles();
  
    customerFieldKeys.forEach((key) => {
      let element;
      const entryValue = entry[key];
  
      if (key === "gender" || key === "nomineeGender") {
        domElements.form.querySelectorAll(`input[name="${key}"]`).forEach((rb) => {
          rb.checked = rb.value === entryValue;
        });
      } else if (key === "partnerName") {
        if (domElements.partnerNameSelect) {
          domElements.partnerNameSelect.value = entryValue || "";
          populateProductDropdown(domElements.partnerNameSelect.value);
        }
      } else if (key === "productName") {
        setTimeout(() => {
          if (domElements.productSelect && domElements.productSelect.disabled === false) {
            domElements.productSelect.value = entryValue || "";
            populatePremiumDropdown(domElements.partnerNameSelect.value, domElements.productSelect.value);
          }
        }, 50);
      } else if (key === "premium") {
        setTimeout(() => {
          if (domElements.premiumSelect && domElements.premiumSelect.disabled === false) {
            const premiumString = entryValue != null ? entryValue.toString() : "";
            domElements.premiumSelect.value = premiumString;
            updateAutoSelectedFields(domElements.partnerNameSelect.value, domElements.productSelect.value, premiumString);
          }
        }, 100);
      } else if (key === "nomineeRelationship") {
        if (domElements.nomineeRelationshipSelect) {
          domElements.nomineeRelationshipSelect.value = entryValue || "";
        }
      } else if (key === "remarks") {
        if (domElements.remarksTextarea) {
          domElements.remarksTextarea.value = entryValue || "";
        }
      }
      else {
        element = document.getElementById(key);
        if (element) {
          element.value = entryValue != null ? entryValue : "";
          // If it's a datepicker, we might need to manually set Flatpickr's date
           if (element.classList.contains('datepicker-input')) {
               const fp = element._flatpickr;
               if (fp) fp.setDate(entryValue, false); // Update without triggering change event
           }
        }
      }
    });
  
    if (domElements.editingIdInput) domElements.editingIdInput.value = id;
    if (domElements.addEntryButton) domElements.addEntryButton.querySelector("span").textContent = translate("btnUpdateEntry");
  
    showToast(translate("toastEditing", { name: entry.customerName || "N/A" }), "info");
    domElements.form.scrollIntoView({ behavior: "smooth", block: "start" });
    applyRequiredAttributes();
}

function handleDeleteEntry(id) {
  const index = collectedData.findIndex((e) => e.id == id);
  if (index > -1) {
    const name = collectedData[index].customerName || "this record";
    if (confirm(translate("DeleteEntry", { name: name }) + ` "${name}"?`)) {
      collectedData.splice(index, 1);
      saveDataToLocalStorage();
      renderTable();
      showToast(translate("toastEntryDeleted", { name: name }), "success");
      if (domElements.editingIdInput.value == id) {
        clearForm();
      }
    }
  } else {
    showToast("Could not find the entry to delete.", "error");
  }
}


// --- Firebase Submission ---
async function handleSubmitToFirebase() {
  if (collectedData.length === 0) {
    showToast(translate("toastNothingToSubmit"), "warning");
    return;
  }
  showLoader();
  
  const batch = writeBatch(db); 

  collectedData.forEach(entry => {
    const docRef = doc(dataCollection);
    const { id, ...dataToSend } = entry;
    
    dataToSend.createdAt = serverTimestamp();
    const derived = getDerivedPartnerData(dataToSend);
    dataToSend.tenure = derived.tenure;
    dataToSend.cseName = derived.cseName;
    
    batch.set(docRef, dataToSend);
  });

  try {
    await batch.commit();
    showToast(translate("toastSubmitSuccess", {count: collectedData.length}), "success");
    clearLocalStorageAndCustomerData();
    uploadedImages = [];
    updateImageViewerState();
  } catch (error) {
    console.error("Error submitting batch to Firestore:", error);
    showToast(translate("toastFirebaseError", { message: error.message }), "error");
  } finally {
    hideLoader();
  }
}

// Helper function to get derived data (to avoid code duplication)
function getDerivedPartnerData(entry) {
    let tenure = "";
    let cseName = "";
    if (entry.partnerName && entry.productName && entry.premium != null) {
        const premiumNum = parseFloat(entry.premium);
        const productDetails = partnerData[entry.partnerName]?.[entry.productName];
        if (Array.isArray(productDetails)) {
            const selectedTier = productDetails.find((tier) => tier.Premium === premiumNum);
            if (selectedTier) {
                tenure = selectedTier.Tenure;
                cseName = selectedTier["CSE Name"];
            }
        }
    }
    return { tenure, cseName };
}


// --- Duplicate Detection ---
function findDuplicates(data) {
  const combinations = new Map();
  const duplicateIds = new Set();
  if (!data || !Array.isArray(data)) {
    return duplicateIds;
  }
  data.forEach((entry) => {
    if (!entry || typeof entry !== "object") return;
    const keyParts = duplicateCheckFields.map((field) =>
      (entry[field] || "").toString().trim().toLowerCase()
    );
    if (keyParts.every((part) => part)) {
      const key = keyParts.join("|");
      if (combinations.has(key)) {
        duplicateIds.add(entry.id);
        duplicateIds.add(combinations.get(key));
      } else {
        combinations.set(key, entry.id);
      }
    }
  });
  return duplicateIds;
}

// --- Table Sorting ---
let sortColumn = null;
let sortDirection = "asc";

function sortData(key) {
  const isAsc = sortColumn === key && sortDirection === "asc";
  sortDirection = isAsc ? "desc" : "asc";
  sortColumn = key;

  const getDerivedValue = (entry, sortKey) => {
    if (sortKey === "tenure" || sortKey === "cseName") {
      if (entry.partnerName && entry.productName && entry.premium != null) {
        const premiumNum = parseFloat(entry.premium);
        const productDetails =
          partnerData[entry.partnerName]?.[entry.productName];
        if (Array.isArray(productDetails)) {
          const selectedTier = productDetails.find(
            (tier) => tier.Premium === premiumNum
          );
          if (selectedTier) {
            return sortKey === "tenure"
              ? selectedTier.Tenure
              : selectedTier["CSE Name"];
          }
        }
      }
      return null;
    }
    return entry[sortKey];
  };

  collectedData.sort((a, b) => {
    let valA = getDerivedValue(a, key);
    let valB = getDerivedValue(b, key);
    const typeA = typeof valA;
    const typeB = typeof valB;
    const nullSortOrder = isAsc ? 1 : -1;
    if (valA == null && valB == null) return 0;
    if (valA == null) return nullSortOrder;
    if (valB == null) return -nullSortOrder;

    if (key === "premium" || key === "tenure") {
      const numA = Number(valA);
      const numB = Number(valB);
      if (!isNaN(numA) && !isNaN(numB)) {
        return isAsc ? numA - numB : numB - numA;
      }
      return isNaN(numA) ? 1 : -1;
    } else if (
      key === "enrolmentDate" ||
      key === "dob" ||
      key === "nomineeDob"
    ) {
      const dateA = valA ? new Date(valA.split("-").reverse().join("-")) : null; // Handles d-m-y format
      const dateB = valB ? new Date(valB.split("-").reverse().join("-")) : null; // Handles d-m-y format
      const dateAValid = dateA && !isNaN(dateA.getTime());
      const dateBValid = dateB && !isNaN(dateB.getTime());
      if (dateAValid && dateBValid)
        return isAsc ? dateA - dateB : dateB - dateA;
      else if (dateAValid) return isAsc ? -1 : 1;
      else if (dateBValid) return isAsc ? 1 : -1;
      else return 0;
    } else {
      const strA = String(valA).toLowerCase();
      const strB = String(valB).toLowerCase();
      return isAsc ? strA.localeCompare(strB) : strB.localeCompare(strA);
    }
  });
  renderTable();
}

function updateSortIcons() {
  document
    .querySelectorAll("#dataTable thead th[data-sort-key]")
    .forEach((th) => {
      const icon = th.querySelector(".sort-icon");
      if (!icon) return;
      if (th.dataset.sortKey === sortColumn) {
        icon.textContent = sortDirection === "asc" ? "▲" : "▼";
        icon.style.opacity = "1";
      } else {
        icon.textContent = "▲";
        icon.style.opacity = "0.3";
      }
    });
}

// --- Table Rendering ---
function renderTable() {
  if (!domElements.tableBody || !domElements.dataTableHead) return;
  domElements.tableBody.innerHTML = "";

  const headerCells = domElements.dataTableHead.querySelectorAll("th");
  const colspan = headerCells.length;

  if (collectedData.length === 0) {
    domElements.tableBody.innerHTML = `<tr class="no-data-row"><td colspan="${colspan}">${translate(
      "tblNoEntries"
    )}</td></tr>`;
    updateSortIcons();
    return;
  }

  let duplicateIds = findDuplicates(collectedData);
  const oneHourAgo = Date.now() - 3600000;
  const fragment = document.createDocumentFragment();

  try {
    collectedData.forEach((entry) => {
      const row = document.createElement("tr");
      row.dataset.id = entry.id;

      let tenure = "";
      let cseName = "";
      if (entry.partnerName && entry.productName && entry.premium != null) {
        const premiumNum = parseFloat(entry.premium);
        const productDetails =
          partnerData[entry.partnerName]?.[entry.productName];
        if (Array.isArray(productDetails)) {
          const selectedTier = productDetails.find(
            (tier) => tier.Premium === premiumNum
          );
          if (selectedTier) {
            tenure = selectedTier.Tenure;
            cseName = selectedTier["CSE Name"];
          }
        }
      }

      let isMissingRequired = criticalFieldsForBlanks.some((key) => {
        if (key === "premium")
          return entry[key] == null || isNaN(parseFloat(entry[key]));
        return !entry[key];
      });
      let isRecent = entry.id > oneHourAgo;
      let isDuplicate = duplicateIds && duplicateIds.has(entry.id);

      row.className = "fade-in";
      if (isDuplicate) row.classList.add("row-highlight-duplicate");
      else if (isMissingRequired) row.classList.add("row-highlight-missing");
      else if (isRecent) row.classList.add("row-highlight-recent");

      const addCell = (content = "", fieldKey = null, align = "left") => {
        const cell = row.insertCell();
        cell.style.textAlign = align;
        const editableFields = [
          "branchName", "region", "customerName", "customerID", "mobileNumber",
          "enrolmentDate", "dob", "savingsAccountNumber", "csbCode", "d2cRoCode",
          "nomineeName", "nomineeRelationship", "nomineeDob", "nomineeMobileNumber", "remarks",
        ];
        if (isTableFullscreen && editableFields.includes(fieldKey)) {
          const isDate = fieldKey === "enrolmentDate" || fieldKey === "dob" || fieldKey === "nomineeDob";
          const inputType = isDate ? "text" : (fieldKey === "mobileNumber" || fieldKey === "nomineeMobileNumber" ? "tel" : "text");
          const inputElementTag = fieldKey === "remarks" ? "textarea" : "input";
          const input = document.createElement(inputElementTag);
          if (inputElementTag === "input") input.type = inputType;
          if (inputElementTag === "textarea") input.rows = 1;
          if (inputType === "tel") input.pattern = "\\d{10,12}";
          input.value = content != null ? content : "";
          input.dataset.id = entry.id;
          input.dataset.key = fieldKey;
          input.classList.add("table-input-edit");
          if(isDate) input.classList.add("datepicker-input"); // Add class for flatpickr
          cell.appendChild(input);
          if(isDate && typeof flatpickr !== 'undefined') { // Initialize flatpickr on the new input
             flatpickr(input, { dateFormat: "d-m-Y", allowInput: true });
          }
        } else if (
          fieldKey === "premium" &&
          content != null &&
          !isNaN(content)
        ) {
          cell.textContent = formatCurrency(Number(content));
        } else {
          cell.textContent = content != null ? String(content) : "";
        }
        if (fieldKey) cell.dataset.key = fieldKey;
        return cell;
      };

      addCell(entry.branchName, "branchName");
      addCell(entry.region, "region"); // NEW
      addCell(entry.customerName, "customerName");
      addCell(entry.customerID, "customerID");
      addCell(translate(`optGender${entry.gender || "Unknown"}`), "gender");
      addCell(entry.mobileNumber, "mobileNumber");
      addCell(entry.enrolmentDate, "enrolmentDate");
      addCell(entry.dob, "dob");
      addCell(entry.savingsAccountNumber, "savingsAccountNumber");
      addCell(entry.csbCode, "csbCode");
      addCell(entry.d2cRoCode, "d2cRoCode");
      addCell(entry.partnerName, "partnerName");
      addCell(entry.productName, "productName");
      addCell(entry.premium, "premium", "right");
      addCell(tenure, "tenure", "right");
      addCell(cseName, "cseName");
      addCell(entry.nomineeName, "nomineeName");
      const relKey = `optRel${entry.nomineeRelationship || "Unknown"}`;
      const translatedRel = translate(relKey);
      addCell(
        translatedRel === relKey && entry.nomineeRelationship
          ? entry.nomineeRelationship
          : translatedRel,
        "nomineeRelationship"
      );
      addCell(entry.nomineeDob, "nomineeDob");
      addCell(entry.nomineeMobileNumber, "nomineeMobileNumber");
      addCell(
        translate(`optGender${entry.nomineeGender || "Unknown"}`),
        "nomineeGender"
      );
      addCell(entry.remarks, "remarks");

      const actionCell = row.insertCell();
      actionCell.classList.add("action-cell");
      actionCell.innerHTML = `
                <button class="edit-btn" data-id="${
                  entry.id
                }" title="${translate("titleEditButton")}" ${
        isTableFullscreen ? "disabled" : ""
      }>
                    <i class="fas fa-edit" aria-hidden="true"></i> <span class="sr-only">Edit</span>
                </button>
                <button class="delete-btn" data-id="${
                  entry.id
                }" title="${translate("titleDeleteButton")}">
                    <i class="fas fa-trash-alt" aria-hidden="true"></i> <span class="sr-only">Delete</span>
                </button>
            `;
      fragment.appendChild(row);
    });
  } catch (loopError) {
    console.error("[Debug] CRITICAL ERROR during table row rendering loop:", loopError);
    showToast("Error rendering table rows. Check console.", "error", 10000);
    domElements.tableBody.innerHTML = `<tr class="no-data-row"><td colspan="${colspan}">Error rendering table.</td></tr>`;
    return;
  }

  domElements.tableBody.appendChild(fragment);
  updateSortIcons();
}

// --- Fullscreen & In-Table Editing ---
function handleFullscreenChange() {
  isTableFullscreen = document.fullscreenElement === domElements.tableContainer;
  updateFullscreenButtonText();
  renderTable();
}

function handleInlineTableEdit(event) {
    const input = event.target;
    if (input && input.classList.contains("table-input-edit") && (event.type === "change" || event.type === "blur")) {
        const id = input.dataset.id;
        const key = input.dataset.key;
        let value = input.value.trim();
        if (id && key) {
            const entryIndex = collectedData.findIndex((e) => e.id == id);
            if (entryIndex > -1) {
                if (collectedData[entryIndex][key] !== value) {
                    collectedData[entryIndex][key] = value;
                    saveDataToLocalStorage();
                    input.style.transition = "background-color 0.3s ease";
                    input.style.backgroundColor = "var(--success-hover-color)";
                    setTimeout(() => {
                        input.style.backgroundColor = "";
                    }, 700);
                }
            }
        }
    }
}

function toggleFullscreen() {
  if (!domElements.tableContainer) return;
  if (!document.fullscreenElement) {
    domElements.tableContainer.requestFullscreen().catch((err) => {
      showToast(
        translate("toastFullscreenErr", { message: err.message }),
        "error"
      );
    });
  } else {
    if (document.exitFullscreen) document.exitFullscreen();
  }
}
function updateFullscreenButtonText() {
  if (!domElements.fullscreenButton) return;
  const span = domElements.fullscreenButton.querySelector("span");
  const icon = domElements.fullscreenButton.querySelector("i");
  if (isTableFullscreen) {
    if (span) span.textContent = translate("btnExitFullscreen");
    if (icon) icon.className = "fas fa-compress";
  } else {
    if (span) span.textContent = translate("btnFullscreen");
    if (icon) icon.className = "fas fa-expand";
  }
}

function clearFormValidationStyles() {
  if (!domElements.form) return;
  domElements.form.querySelectorAll('.input-invalid').forEach(el => {
    el.classList.remove('input-invalid');
  });
}

function clearForm() {
  if (!domElements.form) return;
  try {
    formElementIds.forEach((id) => {
      const element = document.getElementById(id);
      if (element) {
        if (element.classList.contains('datepicker-input') && element._flatpickr) {
            element._flatpickr.clear();
        } else if (element.tagName !== 'SELECT' && element.type !== 'radio' && element.type !== 'checkbox') {
            element.value = '';
        }
      }
    });
    ["gender", "nomineeGender"].forEach((groupName) => {
      const radios = domElements.form.querySelectorAll(
        `input[name="${groupName}"]`
      );
      radios.forEach((radio) => (radio.checked = false));
    });
    if (domElements.partnerNameSelect) {
      domElements.partnerNameSelect.value = "";
      populateProductDropdown("");
    }
    if (domElements.nomineeRelationshipSelect) {
      domElements.nomineeRelationshipSelect.value = "";
    }
    if (domElements.editingIdInput) domElements.editingIdInput.value = "";
    if (domElements.addEntryButton)
      domElements.addEntryButton.querySelector("span").textContent =
        translate("btnAddEntry");
    clearFormValidationStyles();
    applyRequiredAttributes();
  } catch (manualResetError) {
    console.error("[Debug] CRITICAL ERROR during manual form reset:", manualResetError);
    showToast("Error resetting form fields. Check console.", "error", 10000);
    try {
      domElements.form.reset();
    } catch (e) {}
  }
}

// --- Download Functions ---
function handleSaveToExcel() {
    if (typeof window.XLSX === "undefined") {
      showToast(translate("toastLibNotLoaded", { libraryName: "XLSX (SheetJS)" }), "error");
      return;
    }
    if (collectedData.length === 0) {
      showToast("No data to export.", "warning");
      return;
    }
    showLoader();
    setTimeout(() => {
      try {
        const exportData = getExportData();
        if (exportData.length === 0) {
          showToast("No exportable data.", "warning");
          hideLoader();
          return;
        }
        const wb = window.XLSX.utils.book_new();
        const headers = Object.keys(exportData[0]);
        const ws = window.XLSX.utils.json_to_sheet(exportData, { header: headers });
        const columnWidths = headers.map(header => ({ wch: Math.max(header.length, 25) }));
        ws["!cols"] = columnWidths;
        window.XLSX.utils.book_append_sheet(wb, ws, "CustomerData");
        const filename = `CustomerData_${new Date().toISOString().slice(0, 10)}.xlsx`;
        window.XLSX.writeFile(wb, filename);
        
        showToast(translate("toastExcelDownloadOnly"), "success", 5000);

      } catch (error) {
        console.error("Error exporting to Excel:", error);
        showToast(translate("toastExcelErr"), "error");
      } finally {
        hideLoader();
      }
    }, 50);
}

function handleSaveToCSV() {
  if (collectedData.length === 0) {
    showToast("No data to export.", "warning");
    return;
  }
  showLoader();
  setTimeout(() => {
    try {
      const data = getExportData();
      if (data.length === 0) {
        showToast("No exportable data.", "warning");
        hideLoader();
        return;
      }
      const headers = Object.keys(data[0]);
      const csvRows = [
        headers.join(","),
        ...data.map((row) =>
          headers
            .map(
              (header) =>
                `"${(row[header] ?? "").toString().replace(/"/g, '""')}"`
            )
            .join(",")
        ),
      ];
      const csvString = csvRows.join("\r\n");
      triggerDownload(
        csvString,
        `CustomerData_${new Date().toISOString().slice(0, 10)}.csv`,
        "text/csv;charset=utf-8;"
      );
      showToast(translate("toastCsvSuccess"), "success");
    } catch (e) {
      console.error("Error exporting to CSV:", e);
      showToast(translate("toastCsvErr"), "error");
    } finally {
      hideLoader();
    }
  }, 50);
}

function handleSaveToPDF() {
  if (
    typeof window.jspdf === "undefined" ||
    typeof window.jspdf.jsPDF === "undefined"
  ) {
    showToast(translate("toastLibNotLoaded", { libraryName: "jsPDF" }), "error");
    return;
  }
  if (typeof window.jspdf.jsPDF.API?.autoTable === "undefined") {
    showToast(translate("toastLibNotLoaded", { libraryName: "jsPDF-AutoTable plugin" }), "error");
    return;
  }
  if (collectedData.length === 0) {
    showToast("No data to export.", "warning");
    return;
  }
  showLoader();
  setTimeout(() => {
    try {
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({
        orientation: "landscape",
        unit: "pt",
        format: "a4",
      });
      const data = getExportData();
      if (data.length === 0) {
        showToast("No exportable data.", "warning");
        hideLoader();
        return;
      }
      const headers = Object.keys(data[0]);
      const body = data.map((row) =>
        headers.map((header) => row[header] ?? "")
      );
      doc.setFontSize(16);
      doc.text(translate("mainHeading"), doc.internal.pageSize.width / 2, 40, { align: "center" });
      doc.autoTable({
        head: [headers],
        body: body,
        startY: 60,
        theme: "striped",
        headStyles: { fillColor: [41, 128, 185], textColor: 255 },
        styles: { fontSize: 5, cellPadding: 1.5 },
        columnStyles: {
          [headers.indexOf(translate("tblColPremium"))]: { halign: "right" },
          [headers.indexOf(translate("tblColTenure"))]: { halign: "right" },
        },
        didDrawPage: function (hookData) {
          doc.setFontSize(8);
          const pageCount = doc.internal.getNumberOfPages();
          doc.text("Page " + hookData.pageNumber + " of " + pageCount, doc.internal.pageSize.width - hookData.settings.margin.right - 40, doc.internal.pageSize.height - 20);
          doc.text(`Generated: ${new Date().toLocaleDateString()}`, hookData.settings.margin.left, doc.internal.pageSize.height - 20);
        },
      });
      doc.save(`CustomerData_${new Date().toISOString().slice(0, 10)}.pdf`);
      showToast(translate("toastPdfSuccess"), "success");
    } catch (e) {
      console.error("Error exporting to PDF:", e);
      showToast(translate("toastPdfErr"), "error");
    } finally {
      hideLoader();
    }
  }, 50);
}

// MODIFIED: Added region to the export data
function getExportData() {
    return collectedData.map((entry) => {
        const { tenure, cseName } = getDerivedPartnerData(entry);
        const genderDisplay = translate(`optGender${entry.gender || "Unknown"}`);
        const nomineeGenderDisplay = translate(`optGender${entry.nomineeGender || "Unknown"}`);
        const relKey = `optRel${entry.nomineeRelationship || "Unknown"}`;
        let nomineeRelationshipDisplay = translate(relKey);
        if (nomineeRelationshipDisplay === relKey && entry.nomineeRelationship) {
          nomineeRelationshipDisplay = entry.nomineeRelationship;
        } else if (nomineeRelationshipDisplay === relKey && !entry.nomineeRelationship) {
          nomineeRelationshipDisplay = "";
        }
    
        let rowData = {};
        rowData[translate("tblColBranch")] = entry.branchName || "";
        rowData[translate("tblColRegion")] = entry.region || ""; // NEW
        rowData[translate("tblColName")] = entry.customerName || "";
        rowData[translate("tblColCustID")] = entry.customerID || "";
        rowData[translate("tblColGender")] = genderDisplay;
        rowData[translate("tblColMobile")] = entry.mobileNumber || "";
        rowData[translate("tblColEnrolDate")] = entry.enrolmentDate || "";
        rowData[translate("tblColDOB")] = entry.dob || "";
        rowData[translate("tblColSavingsAcc")] = entry.savingsAccountNumber || "";
        rowData[translate("tblColCSB")] = entry.csbCode || "";
        rowData[translate("tblColD2CRoCode")] = entry.d2cRoCode || "";
        rowData[translate("tblColPartnerName")] = entry.partnerName || "";
        rowData[translate("tblColProductName")] = entry.productName || "";
        rowData[translate("tblColPremium")] = entry.premium != null ? entry.premium : "";
        rowData[translate("tblColTenure")] = tenure;
        rowData[translate("tblColCseName")] = cseName;
        rowData[translate("tblColNomName")] = entry.nomineeName || "";
        rowData[translate("tblColNomRelation")] = nomineeRelationshipDisplay;
        rowData[translate("tblColNomDOB")] = entry.nomineeDob || "";
        rowData[translate("tblColNomMobile")] = entry.nomineeMobileNumber || "";
        rowData[translate("tblColNomGender")] = nomineeGenderDisplay;
        rowData[translate("tblColRemarks")] = entry.remarks || "";
        return rowData;
      });
}

function triggerDownload(content, filename, contentType) {
    const blob = new Blob([content], { type: contentType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// --- Dashboard Functionality (omitted for brevity, no changes were requested here) ---
// ... (The entire dashboard section remains unchanged)

// --- Event Listeners Setup ---
// MODIFIED: Added listeners for new button and live validation
function setupEventListeners() {
    if (domElements.languageSwitcher) domElements.languageSwitcher.addEventListener("change", (e) => setLanguage(e.target.value));
    if (domElements.themeToggleButton) domElements.themeToggleButton.addEventListener("click", toggleTheme);
    if (domElements.toggleDataEntryModeButton) domElements.toggleDataEntryModeButton.addEventListener("click", toggleDataEntryMode);
    if (domElements.fullscreenButton) domElements.fullscreenButton.addEventListener("click", toggleFullscreen);
    if (domElements.submitToFirebaseButton) domElements.submitToFirebaseButton.addEventListener("click", handleSubmitToFirebase);
  
    if (domElements.form) {
      domElements.form.addEventListener("reset", () => setTimeout(clearForm, 0));
      if (domElements.partnerNameSelect) domElements.partnerNameSelect.addEventListener("change", (e) => populateProductDropdown(e.target.value));
      if (domElements.productSelect) domElements.productSelect.addEventListener("change", (e) => populatePremiumDropdown(domElements.partnerNameSelect.value, e.target.value));
      if (domElements.premiumSelect) domElements.premiumSelect.addEventListener("change", (e) => updateAutoSelectedFields(domElements.partnerNameSelect.value, domElements.productSelect.value, e.target.value));
      document.querySelectorAll(".mic-button").forEach((b) => b.addEventListener("click", handleMicButtonClick));
      document.querySelectorAll(".clean-button").forEach((b) => b.addEventListener("click", handleCleanButtonClick));
      if (domElements.addEntryButton) domElements.addEntryButton.addEventListener("click", handleAddOrUpdateEntry);
      if (domElements.clearInputsButton) domElements.clearInputsButton.addEventListener("click", handleCleanAllInputs); // NEW

      // NEW: Live validation listeners
      domElements.form.addEventListener('input', (event) => {
        const target = event.target;
        if(target.hasAttribute('required')) {
            target.classList.toggle('input-invalid', !target.value.trim());
        }
      });
      domElements.form.addEventListener('blur', (event) => {
        const target = event.target;
        if(target.hasAttribute('required')) {
            target.classList.toggle('input-invalid', !target.value.trim());
        }
      }, true); // Use capture to catch blur on all elements
    }
  
    if (domElements.uploadImageButton) domElements.uploadImageButton.addEventListener("click", handleImageUploadTrigger);
    if (domElements.imageUploadInput) domElements.imageUploadInput.addEventListener("change", handleImageFiles);
    if (domElements.toggleImageViewerButton) domElements.toggleImageViewerButton.addEventListener("click", toggleImageViewer);
    if (domElements.prevImageButton) domElements.prevImageButton.addEventListener("click", showPrevImage);
    if (domElements.nextImageButton) domElements.nextImageButton.addEventListener("click", showNextImage);
  
    if (domElements.dataTableHead) domElements.dataTableHead.addEventListener("click", (event) => {
      const header = event.target.closest("th[data-sort-key]");
      if (header) sortData(header.dataset.sortKey);
    });
    if (domElements.tableBody) {
      domElements.tableBody.addEventListener("click", (e) => {
        const button = e.target.closest("button");
        if (!button) return;
        if (button.classList.contains("edit-btn")) handleEditEntry(button.dataset.id);
        else if (button.classList.contains("delete-btn")) handleDeleteEntry(button.dataset.id);
      });
      domElements.tableBody.addEventListener("change", handleInlineTableEdit);
      domElements.tableBody.addEventListener("blur", handleInlineTableEdit, true);
    }
  
    if (domElements.saveExcelButton) domElements.saveExcelButton.addEventListener("click", handleSaveToExcel);
    if (domElements.saveCsvButton) domElements.saveCsvButton.addEventListener("click", handleSaveToCSV);
    if (domElements.savePdfButton) domElements.savePdfButton.addEventListener("click", handleSaveToPDF);
    if (domElements.dashboardButton) domElements.dashboardButton.addEventListener("click", generateDashboard);
    if (domElements.closeDashboardButton) domElements.closeDashboardButton.addEventListener("click", () => {
        if (domElements.dashboardModal) domElements.dashboardModal.style.display = "none";
    });
    if (domElements.prevChartButton) domElements.prevChartButton.addEventListener("click", prevSlide);
    if (domElements.nextChartButton) domElements.nextChartButton.addEventListener("click", nextSlide);
    if (domElements.adminButton) domElements.adminButton.addEventListener("click", handleAdminButtonClick);
    if (domElements.closeAdminLoginButton) domElements.closeAdminLoginButton.addEventListener("click", () => {
        if (domElements.adminLoginModal) domElements.adminLoginModal.style.display = "none";
    });
    if (domElements.adminLoginForm) domElements.adminLoginForm.addEventListener("submit", handleAdminLogin);
    if (domElements.closeAdminEditButton) domElements.closeAdminEditButton.addEventListener("click", () => {
        if (domElements.adminEditModal) domElements.adminEditModal.style.display = "none";
    });
    if (domElements.adminEditForm) domElements.adminEditForm.addEventListener("submit", handleAdminSaveChanges);
    document.addEventListener("fullscreenchange", handleFullscreenChange);
    if (domElements.recordAllButton) domElements.recordAllButton.addEventListener("click", handleRecordAll);
    if (domElements.clearTableButton) domElements.clearTableButton.addEventListener("click", handleClearTable);
    if (domElements.resetApplicationDataButton) domElements.resetApplicationDataButton.addEventListener('click', () => {
        if (confirm("Are you sure you want to reset all application data?")) {
            localStorage.clear();
            window.location.reload();
        }
    });
}

// ... (All other functions like Record All, Admin panel, etc., remain unchanged)
// The existing functions for Dashboard, Admin, etc., are fine as they were.
// I'm keeping them here to ensure you have the full, complete file.

function generateDashboard() {
  if (typeof Chart === "undefined") {
    showToast(
      translate("toastLibNotLoaded", { libraryName: "Chart.js" }),
      "error"
    );
    return;
  }
  if (collectedData.length === 0) {
    showToast("No data available for dashboard.", "warning");
    return;
  }
  showLoader();
  setTimeout(() => {
    try {
      if (genderChartInstance) genderChartInstance.destroy();
      if (branchChartInstance) branchChartInstance.destroy();
      if (partnerChartInstance) partnerChartInstance.destroy();

      const genderCounts = collectedData.reduce((acc, entry) => {
        const genderKey = `optGender${entry.gender || "Unknown"}`;
        acc[genderKey] = (acc[genderKey] || 0) + 1;
        return acc;
      }, {});
      renderGenderChart(genderCounts);

      const branchCounts = collectedData.reduce((acc, entry) => {
        const branchName = entry.branchName?.trim() || "Unknown Branch";
        acc[branchName] = (acc[branchName] || 0) + 1;
        return acc;
      }, {});
      const sortedBranches = Object.entries(branchCounts)
        .sort(([, a], [, b]) => b - a)
        .slice(0, 5);
      renderBranchChart(Object.fromEntries(sortedBranches));

      const partnerCounts = collectedData.reduce((acc, entry) => {
        const partnerName = entry.partnerName || "No Partner Selected";
        acc[partnerName] = (acc[partnerName] || 0) + 1;
        return acc;
      }, {});
      renderPartnerChart(partnerCounts);

      const totalEntries = collectedData.length;
      const femaleCount = genderCounts["optGenderFemale"] || 0;
      const femaleRatio =
        totalEntries > 0 ? ((femaleCount / totalEntries) * 100).toFixed(1) : 0;
      const blankEntries = collectedData.filter((entry) =>
        criticalFieldsForBlanks.some((key) => {
          if (key === "premium")
            return entry[key] == null || isNaN(parseFloat(entry[key]));
          return !entry[key];
        })
      ).length;
      document.getElementById("statTotalEntries").textContent = totalEntries;
      document.getElementById(
        "statFemaleRatio"
      ).textContent = `${femaleRatio}%`;
      document.getElementById("statBlankEntries").textContent = blankEntries;

      currentChartSlide = 0;
      showSlide(currentChartSlide);
      if (domElements.dashboardModal) {
        domElements.dashboardModal.style.display = "block";
        requestAnimationFrame(() => {
          domElements.dashboardModal.classList.add("fade-in-modal");
        });
      }
    } catch (e) {
      console.error("Error generating dashboard:", e);
      showToast(translate("toastDashErr"), "error");
    } finally {
      hideLoader();
    }
  }, 50);
}
function renderGenderChart(data) {
  const ctx = document.getElementById("genderChart")?.getContext("2d");
  if (!ctx) return;
  const translatedLabels = Object.keys(data).map((key) => translate(key));
  const chartData = Object.values(data);
  const backgroundColors = [
    "rgba(54, 162, 235, 0.7)",
    "rgba(255, 99, 132, 0.7)",
    "rgba(255, 206, 86, 0.7)",
    "rgba(75, 192, 192, 0.7)",
    "rgba(153, 102, 255, 0.7)",
    "rgba(255, 159, 64, 0.7)",
  ];
  genderChartInstance = new Chart(ctx, {
    type: "pie",
    data: {
      labels: translatedLabels,
      datasets: [
        {
          label: translate("DashboardGenderTitle"),
          data: chartData,
          backgroundColor: backgroundColors.slice(0, chartData.length),
          borderColor: backgroundColors
            .map((c) => c.replace("0.7", "1"))
            .slice(0, chartData.length),
          borderWidth: 1,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: { display: true, text: translate("DashboardGenderTitle") },
        legend: { position: "top" },
      },
    },
  });
}
function renderBranchChart(data) {
  const ctx = document.getElementById("branchChart")?.getContext("2d");
  if (!ctx) return;
  const labels = Object.keys(data);
  const chartData = Object.values(data);
  const bgColor = "rgba(75, 192, 192, 0.7)";
  branchChartInstance = new Chart(ctx, {
    type: "bar",
    data: {
      labels: labels,
      datasets: [
        {
          label: "Entries",
          data: chartData,
          backgroundColor: bgColor,
          borderColor: bgColor.replace("0.7", "1"),
          borderWidth: 1,
        },
      ],
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: { display: true, text: translate("DashboardBranchTitle") },
        legend: { display: false },
      },
      scales: { x: { beginAtZero: true } },
    },
  });
}
function renderPartnerChart(data) {
  const ctx = document.getElementById("partnerChart")?.getContext("2d");
  if (!ctx) return;
  const labels = Object.keys(data);
  const chartData = Object.values(data);
  const backgroundColors = [
    "rgba(255, 159, 64, 0.7)",
    "rgba(153, 102, 255, 0.7)",
    "rgba(75, 192, 192, 0.7)",
    "rgba(255, 99, 132, 0.7)",
    "rgba(54, 162, 235, 0.7)",
    "rgba(255, 206, 86, 0.7)",
  ];
  partnerChartInstance = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: labels,
      datasets: [
        {
          label: translate("DashboardPartnerTitle"),
          data: chartData,
          backgroundColor: backgroundColors.slice(0, chartData.length),
          borderColor: backgroundColors
            .map((c) => c.replace("0.7", "1"))
            .slice(0, chartData.length),
          borderWidth: 1,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: { display: true, text: translate("DashboardPartnerTitle") },
        legend: { position: "top" },
      },
    },
  });
}
function showSlide(index) {
  if (!domElements.dashboardCarouselSlides) return;
  const slides =
    domElements.dashboardCarouselSlides.querySelectorAll(".chart-slide");
  const totalSlides = slides.length;
  if (totalSlides === 0) return;
  if (index >= totalSlides) index = totalSlides - 1;
  if (index < 0) index = 0;
  currentChartSlide = index;
  const offset = -index * 100;
  domElements.dashboardCarouselSlides.style.transform = `translateX(${offset}%)`;
  if (domElements.prevChartButton)
    domElements.prevChartButton.disabled = index === 0;
  if (domElements.nextChartButton)
    domElements.nextChartButton.disabled = index === totalSlides - 1;
}
function nextSlide() {
  showSlide(currentChartSlide + 1);
}
function prevSlide() {
  showSlide(currentChartSlide - 1);
}
function handleAdminButtonClick() {
  if (!domElements.adminLoginModal || !domElements.adminLoginForm) return;
  domElements.adminLoginForm.reset();
  domElements.adminLoginModal.style.display = "block";
  requestAnimationFrame(() => {
    domElements.adminLoginModal.classList.add("fade-in-modal");
  });
  if (domElements.adminUsernameInput) domElements.adminUsernameInput.focus();
}
function handleAdminLogin(event) {
  event.preventDefault();
  if (
    !domElements.adminLoginForm ||
    !domElements.adminLoginModal ||
    !domElements.adminEditModal ||
    !domElements.adminUsernameInput ||
    !domElements.adminPasswordInput
  )
    return;
  const username = domElements.adminUsernameInput.value;
  const password = domElements.adminPasswordInput.value;
  if (
    username.toLowerCase() === ADMIN_USERNAME.toLowerCase() &&
    password.toLowerCase() === ADMIN_PASSWORD.toLowerCase()
  ) {
    domElements.adminLoginModal.style.display = "none";
    domElements.adminLoginModal.classList.remove("fade-in-modal");
    populateAdminEditPanel();
    domElements.adminEditModal.style.display = "block";
    requestAnimationFrame(() => {
      domElements.adminEditModal.classList.add("fade-in-modal");
    });
    showToast(translate("toastLoginSuccess"), "success");
  } else {
    showToast(translate("toastLoginFail"), "error");
    domElements.adminLoginForm.reset();
    domElements.adminUsernameInput.focus();
  }
}
function populateAdminEditPanel() {
  if (!domElements.adminEditGrid || !partnerData) {
    if (domElements.adminEditGrid)
      domElements.adminEditGrid.innerHTML = `<p>${translate(
        "toastAdminLoadError"
      )}</p>`;
    return;
  }
  domElements.adminEditGrid.innerHTML = "";
  const sortedPartnerNames = Object.keys(partnerData).sort();
  sortedPartnerNames.forEach((partnerName) => {
    const partnerDetails = partnerData[partnerName];
    const partnerGroupDiv = document.createElement("div");
    partnerGroupDiv.className = "admin-partner-group";
    partnerGroupDiv.dataset.partnerName = partnerName;
    const partnerTitle = document.createElement("h4");
    partnerTitle.textContent = partnerName;
    partnerGroupDiv.appendChild(partnerTitle);
    const sortedProductNames = Object.keys(partnerDetails).sort();
    sortedProductNames.forEach((productName) => {
      const productTiers = partnerDetails[productName];
      const productGroupDiv = document.createElement("div");
      productGroupDiv.className = "admin-product-group";
      productGroupDiv.dataset.productName = productName;
      const productHeaderDiv = document.createElement("div");
      productHeaderDiv.style.display = "flex";
      productHeaderDiv.style.justifyContent = "space-between";
      productHeaderDiv.style.alignItems = "center";
      productHeaderDiv.style.marginBottom = "10px";
      const productLabel = document.createElement("h5");
      productLabel.textContent = productName;
      productHeaderDiv.appendChild(productLabel);
      productGroupDiv.appendChild(productHeaderDiv);
      if (Array.isArray(productTiers)) {
        const sortedTiers = [...productTiers].sort(
          (a, b) => a.Premium - b.Premium
        );
        const originalIndices = sortedTiers.map((st) =>
          partnerDetails[productName].findIndex(
            (ot) =>
              ot.Premium === st.Premium &&
              ot.Tenure === st.Tenure &&
              ot["CSE Name"] === st["CSE Name"]
          )
        );
        sortedTiers.forEach((tier, displayIndex) => {
          const originalIndex = originalIndices[displayIndex];
          const tierGroupDiv = document.createElement("div");
          tierGroupDiv.className = "admin-tier-group";
          tierGroupDiv.dataset.tierIndex = originalIndex;
          const premiumField = document.createElement("div");
          premiumField.className = "form-field admin-field";
          premiumField.innerHTML = `
                        <label for="admin_prem_${partnerName}_${productName}_${originalIndex}" data-translate-key="titleAdminPremium">${translate(
            "titleAdminPremium"
          )}:</label>
                        <input type="number" id="admin_prem_${partnerName}_${productName}_${originalIndex}" value="${
            tier.Premium
          }" required min="0" step="any" data-field="Premium">
                    `;
          tierGroupDiv.appendChild(premiumField);
          const tenureField = document.createElement("div");
          tenureField.className = "form-field admin-field";
          tenureField.innerHTML = `
                        <label for="admin_ten_${partnerName}_${productName}_${originalIndex}" data-translate-key="titleAdminTenure">${translate(
            "titleAdminTenure"
          )}:</label>
                        <input type="number" id="admin_ten_${partnerName}_${productName}_${originalIndex}" value="${
            tier.Tenure
          }" required min="0" step="1" data-field="Tenure">
                    `;
          tierGroupDiv.appendChild(tenureField);
          const cseField = document.createElement("div");
          cseField.className = "form-field admin-field";
          cseField.innerHTML = `
                        <label for="admin_cse_${partnerName}_${productName}_${originalIndex}" data-translate-key="titleAdminCseName">${translate(
            "titleAdminCseName"
          )}:</label>
                        <input type="text" id="admin_cse_${partnerName}_${productName}_${originalIndex}" value="${
            tier["CSE Name"]
          }" required data-field="CSE Name">
                     `;
          tierGroupDiv.appendChild(cseField);
          productGroupDiv.appendChild(tierGroupDiv);
        });
      }
      partnerGroupDiv.appendChild(productGroupDiv);
    });
    domElements.adminEditGrid.appendChild(partnerGroupDiv);
  });
}
function handleAdminSaveChanges(event) {
  event.preventDefault();
  if (!domElements.adminEditGrid) return;
  let hasChanges = false;
  let isValid = true;
  const updatedPartnerData = JSON.parse(JSON.stringify(partnerData));
  domElements.adminEditGrid
    .querySelectorAll(".admin-partner-group")
    .forEach((partnerGroup) => {
      if (!isValid) return;
      const partnerName = partnerGroup.dataset.partnerName;
      partnerGroup
        .querySelectorAll(".admin-product-group")
        .forEach((productGroup) => {
          if (!isValid) return;
          const productName = productGroup.dataset.productName;
          productGroup
            .querySelectorAll(".admin-tier-group")
            .forEach((tierGroup) => {
              if (!isValid) return;
              const tierIndex = parseInt(tierGroup.dataset.tierIndex, 10);
              if (
                isNaN(tierIndex) ||
                !updatedPartnerData[partnerName]?.[productName]?.[tierIndex]
              ) {
                isValid = false;
                return;
              }
              let currentTier =
                updatedPartnerData[partnerName][productName][tierIndex];
              let tierChanged = false;
              tierGroup
                .querySelectorAll("input[data-field]")
                .forEach((input) => {
                  if (!isValid) return;
                  const field = input.dataset.field;
                  const value = input.value.trim();
                  let parsedValue = value;
                  input.style.borderColor = "";
                  if (field === "Premium" || field === "Tenure") {
                    parsedValue = parseFloat(value);
                    if (isNaN(parsedValue) || parsedValue < 0) {
                      input.style.borderColor = "var(--danger-color)";
                      isValid = false;
                    } else {
                      if (Number(currentTier[field]) !== parsedValue)
                        tierChanged = true;
                    }
                  } else if (field === "CSE Name") {
                    parsedValue = value;
                    if (!parsedValue) {
                      input.style.borderColor = "var(--danger-color)";
                      isValid = false;
                    } else {
                      if (currentTier[field] !== parsedValue)
                        tierChanged = true;
                    }
                  } else {
                    return;
                  }
                  if (isValid && tierChanged) {
                    currentTier[field] = parsedValue;
                  }
                });
              if (tierChanged) hasChanges = true;
            });
        });
    });
  if (!isValid) {
    showToast(translate("toastAdminSaveValidation"), "warning");
    const firstInvalid = domElements.adminEditGrid.querySelector(
      'input[style*="border-color"]'
    );
    if (firstInvalid) firstInvalid.focus();
    return;
  }
  if (hasChanges) {
    partnerData = updatedPartnerData;
    savePartnerData();
    populatePartnerNameDropdown();
    renderTable();
    showToast(translate("toastAdminSaveSuccess"), "success");
  } else {
    showToast(translate("toastAdminSaveNoChange"), "info");
  }
  if (domElements.adminEditModal) {
    domElements.adminEditModal.style.display = "none";
    domElements.adminEditModal.classList.remove("fade-in-modal");
  }
}
function enableAllMicButtons() {
  if (!isRecordingAllActive) {
    document.querySelectorAll(".mic-button").forEach((btn) => {
      btn.disabled = false;
      btn.classList.remove("listening");
    });
  }
}
function disableAllMicButtons() {
  document.querySelectorAll(".mic-button").forEach((btn) => {
    btn.disabled = true;
    btn.classList.remove("listening");
  });
}
function handleRecordAll() {
  if (!recognition) {
    showToast(translate("toastSpeechUnsupp"), "error");
    if (domElements.recordAllButton)
      domElements.recordAllButton.disabled = true; // Disable if speech not supported
    return;
  }
  if (isRecordingAllActive) {
    stopRecordAllSequence("Recording cancelled by user.");
  } else {
    isRecordingAllActive = true;
    currentRecordAllIndex = 0;
    if (domElements.recordAllButton) {
      let cancelText = "Cancel Recording";
      if (translations[currentLang]?.btnCancelRecordAll) {
        cancelText = translate("btnCancelRecordAll");
      } else if (translations.en?.btnCancelRecordAll) {
        cancelText = translations.en.btnCancelRecordAll;
      }
      domElements.recordAllButton.querySelector("span").textContent =
        cancelText;
      domElements.recordAllButton.title = cancelText;
      domElements.recordAllButton.classList.add("recording-active");
    }
    disableAllMicButtons();
    if (!recordAllStatusElement && domElements.form) {
      recordAllStatusElement = document.createElement("div");
      recordAllStatusElement.id = "record-all-status";
      recordAllStatusElement.style.marginTop = "10px";
      recordAllStatusElement.style.fontWeight = "bold";
      const formParent = domElements.form.parentNode;
      if (formParent) {
        formParent.insertBefore(
          recordAllStatusElement,
          domElements.form.nextSibling
        );
      } else {
        document.body.appendChild(recordAllStatusElement);
      }
    }
    updateRecordAllStatus("Starting Record All...", false);
    processNextRecordAllField();
  }
}
function processNextRecordAllField() {
  clearRecordAllStatus();
  if (
    !isRecordingAllActive ||
    currentRecordAllIndex >= recordAllSequence.length
  ) {
    stopRecordAllSequence(
      currentRecordAllIndex >= recordAllSequence.length
        ? "Recording sequence complete."
        : "Recording stopped."
    );
    return;
  }
  const fieldId = recordAllSequence[currentRecordAllIndex];
  const inputElement = document.getElementById(fieldId);
  if (!inputElement) {
    currentRecordAllIndex++;
    setTimeout(processNextRecordAllField, 50);
    return;
  }
  const isVisible =
    inputElement.offsetWidth > 0 ||
    inputElement.offsetHeight > 0 ||
    inputElement.getClientRects().length > 0;
  if (!isVisible || inputElement.disabled || inputElement.readOnly) {
    currentRecordAllIndex++;
    setTimeout(processNextRecordAllField, 50);
    return;
  }
  currentTargetInput = inputElement;
  inputElement.focus();
  applyFieldHighlight(fieldId);
  const labelElement = domElements.form?.querySelector(
    `label[for='${fieldId}']`
  );
  let fieldLabel = fieldId;
  if (labelElement) {
    const labelClone = labelElement.cloneNode(true);
    labelClone
      .querySelectorAll(".required-star, .sr-only")
      .forEach((s) => s.remove());
    fieldLabel = labelClone.textContent.replace(/[:*]/g, "").trim() || fieldId;
  }
  updateRecordAllStatus(`Listening for: ${fieldLabel}...`, false);
  try {
    recognition.lang = domElements.languageSelect?.value || "en-US";
    recognition.start();
  } catch (error) {
    removeFieldHighlight(fieldId);
    currentRecordAllIndex++;
    setTimeout(processNextRecordAllField, 100);
  }
}
function stopRecordAllSequence(message = "Recording stopped.") {
  isRecordingAllActive = false;
  currentRecordAllIndex = -1;
  if (currentTargetInput) removeFieldHighlight(currentTargetInput.id);
  currentTargetInput = null;
  if (domElements.recordAllButton) {
    domElements.recordAllButton.querySelector("span").textContent =
      translate("lblRecordAll");
    domElements.recordAllButton.title = translate("btnRecordAll");
    domElements.recordAllButton.classList.remove("recording-active");
  }
  updateRecordAllStatus(message, true);
  setTimeout(clearRecordAllStatus, 4000);
  enableAllMicButtons();
}
function applyFieldHighlight(fieldId) {
  if (!fieldId) return;
  const element = document.getElementById(fieldId);
  if (element) {
    document
      .querySelectorAll(".recording-active-field")
      .forEach((el) => el.classList.remove("recording-active-field"));
    element.classList.add("recording-active-field");
  }
}
function removeFieldHighlight(fieldId) {
  if (!fieldId) return;
  const element = document.getElementById(fieldId);
  if (element) element.classList.remove("recording-active-field");
}
function updateRecordAllStatus(text, isFinal = false) {
  if (recordAllStatusElement) {
    recordAllStatusElement.textContent = text;
    recordAllStatusElement.style.color = isFinal
      ? "var(--success-color)"
      : "var(--primary-color)";
  }
}
function clearRecordAllStatus() {
  if (recordAllStatusElement) recordAllStatusElement.textContent = "";
}
function handleClearTable() {
    if (confirm(translate("DeleteEntry", { name: "ALL entries" }).replace("the entry for", "ALL entries") + "? This will clear the local table and stored data.")) {
        clearLocalStorageAndCustomerData();
        if (uploadedImages && uploadedImages.length > 0) {
            uploadedImages = [];
            updateImageViewerState();
        }
    }
}