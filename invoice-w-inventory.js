const XlsxPopulate = require("xlsx-populate");
const XLSX = require("xlsx");
const moment = require("moment");
const fs = require("fs");

// company infor
const companyInfo = {
  addressLine1: "Paya Ubi Industrial Estatev tesstttt tesssttt",
  addressLine2: "#01-122222 66666777",
  postalCode: "2289100099",
  country: "Singaporeee",
  telNo: "678810011111666",
  faxNo: "689012231277",
  email: "sales122@doorsuisse.sggg",
  website: "https://doorsuisse.sg/",
};

// invoice general
const invoice = {
  invoiceNo: "INV4180",
  projectNo: "PM5332",
  createdDate: "2023-06-16T04:19:42.993Z",
  estimateStartDate: "2023-06-16T04:19:42.993Z",
  contractPeriod: {
    from: "2023-06-04T17:00:00.000Z",
    to: "2023-06-22T17:00:00.000Z",
  },
  gst: 8,
  discount: 0,
};

// customer detail
const customer = {
  isActive: true,
  isDeleted: false,
  _id: "6486c6a1cc7799031c62bd7a",
  id: "f4abacd9-05f8-4fb6-882d-ea195c898aa7",
  createdDate: "2023-06-12T07:17:53.105Z",
  updatedDate: "2023-06-12T07:17:53.105Z",
  customerType: "Corporate",
  uenNo: "uen276",
  picName: "ngochieu276",
  picEmail: "ngochieu276@gmail,.com",
  pocName: "ngochieu",
  pocEmail: "ngochieuPOC",
  company: "ngochieu",
  picContactNo: {
    ext: "+65",
    phoneNumber: "6531313123",
  },
  pocContactNo: {
    ext: "+65",
    phoneNumber: "6512312313",
  },
  billingAddress: {
    addressLine1: "adress 1",
    addressLine2: "address 2",
    postalCode: "100000",
  },
  remarks: "customer reakadsf",
};
// billing address
const billingAddress = {
  addressLine1: "adress 1",
  addressLine2: "address 2",
  postalCode: "100000",
};
// locations
const projectLocation = [
  {
    addressLine1: "address 1",
    addressLine2: "address 2",
    postalCode: "100000",
    projectLocationName: "project location1",
    remarks: "asdfasdf",
  },
  {
    addressLine1: "address 3",
    addressLine2: "address 4",
    postalCode: "100000",
    projectLocationName: "project location2",
    remarks: "asdfasdf",
  },
];
// doorType
const doorTypes = [
  {
    doorTypeId: "a896cc34-95af-402f-9e6f-371e5a8aff6a",
    workOrderNo: "WO9328",
    doorId: "7751d594-67e9-453b-8e82-0fdecbf68db0",
    doorNo: "PM2299-1",
    estCompletedDate: "2023-06-11T17:00:00.000Z",
    door: {
      doorTypeIndex: 1,
      doorType: "Single",
      doorMaterial: "Timber",
      floorGuideType: "U-15",
      doorTypeOtherText: "",
      doorMaterialOtherText: "",
      floorGuideTypeOtherText: "",
      doorSize: {
        height: 1102,
        length: 901,
        thickness: 153,
        trackLength: 1104,
        coverLength: 1105,
        glassClampSize: 116,
      },
      remarks: "remaks",
    },
  },
  {
    doorTypeId: "7c434876-e548-465f-9bc8-8beeb3b926d7",
    workOrderNo: "WO5832",
    doorId: "0797b644-5acb-4417-8475-c1ea6687a605",
    doorNo: "PM2299-2",
    estCompletedDate: "2023-06-21T17:00:00.000Z",
    door: {
      doorTypeIndex: 1,
      doorType: "Others",
      doorMaterial: "Timber",
      floorGuideType: "U-15",
      doorTypeOtherText: "fdasdf",
      doorMaterialOtherText: "",
      floorGuideTypeOtherText: "",
      doorSize: {
        height: 1104,
        length: 905,
        thickness: 153,
        trackLength: 1102,
        coverLength: 1101,
        glassClampSize: 110,
      },
      remarks: "adf",
    },
  },
];
// inventories
const inventories = [
  {
    doorType: "7de7671c-c4fe-465d-98c0-edba5f4fc8b2",
    inventoryId: "3e163267-725a-42cb-8d25-9948096f55ab",
    categoryName: "Category 2345555",
    locationName: "Location 1",
    inventoryName: "Inve11",
    sku: "fg5555",
    qty: 0,
  },
  {
    doorType: "7de7671c-c4fe-465d-98c0-edba5f4fc8b2",
    inventoryId: "779b0865-90a2-4cc3-853b-f51a561805ea",
    categoryName: "Category 1",
    locationName: "Location 1",
    inventoryName: "Test ACB",
    sku: "SKU CH777",
    qty: 0,
  },
];
// products
const products = [
  {
    productId: "83ac1d73-613b-43a8-98ac-d6ef57fef8b2",
    qty: 1,
    productName: "33334442222",
    unit: "Set",
    price: 100,
  },
  {
    productId: "1c306569-2e7b-4af7-b984-2d6247583db2",
    qty: 1,
    productName: "123123",
    unit: "Set",
    price: 122,
  },
];
// remarks
const remarks = "remarkssssssssss";
// terms
const terms = "remarkssssssssss";
// warranty
const warranty = {
  pic: "pic123",
  picEmail: "picemail@gmail.com",
  location: "location 1234",
  startDate: "2023-05-29T17:00:00.000Z",
  endDate: "2023-05-29T17:00:00.000Z",
};
// payment term
const invoicePaymentTerm = "Cash";
// Sub invoice detail
const subInvoice = {
  invoiceNo: "INV0001",
  dateIssued: "2023-05-29T17:00:00.000Z",
};
// Sub invoice products
const subInvoiceProductServices = [
  {
    productId: "3dfa769b-47f1-48be-be38-cc1dcffd0047",
    qty: 1,
    productName: "Product4",
    unit: "Set",
    price: 15,
  },
  {
    productId: "13a7915c-038e-4faf-a5e1-9fe2f4c7fe81",
    qty: 1,
    productName: "product 3",
    unit: "Set",
    price: 500,
  },
];
// Sub invoice payment term
const subInvoicePaymentTerm = "Credit Terms";

const renderLineSpace = (lines) => {
  let emptyArray = [];
  for (let i = 0; i < lines; i++) {
    emptyArray.push("");
  }
  return emptyArray;
};

const fileName = `${invoice.invoiceNo}-with-inventories.xlsx`;

const handleExport = ({
  companyInfo,
  invoice,
  customer,
  billingAddress,
  projectLocation,
  doorTypes,
  inventories,
  products,
  remarks,
  terms,
  warranty,
  invoicePaymentTerm,
  subInvoice,
  subInvoiceProductServices,
  subInvoicePaymentTerm,
}) => {
  const space0 = 2;

  let companyInforTable = [];

  companyInforTable.push({
    A: "Address:",
    B: companyInfo.addressLine1,
    C: companyInfo.addressLine2,
  });
  companyInforTable.push({
    A: "Postal Code:",
    B: companyInfo.postalCode,
  });
  companyInforTable.push({
    A: "Tel No.",
    B: companyInfo.telNo,
  });
  companyInforTable.push({
    A: "Fax No.",
    B: companyInfo.faxNo,
  });
  companyInforTable.push({
    A: "Email:",
    B: companyInfo.email,
  });
  companyInforTable.push({
    A: "Website:",
    B: companyInfo.website,
  });

  const companyInforLines = companyInforTable.length;
  const spaceAfterCompanyInfor = 4;

  let invoiceTable = [];
  invoiceTable.push({
    A: "Invoice",
  });
  invoiceTable.push({});
  invoiceTable.push({
    A: "Project No.:",
    B: invoice.projectNo,
  });
  invoiceTable.push({
    A: "Estimated Project Start Date:",
    B: `${invoice.createdDate} check!`,
  });
  invoiceTable.push({
    A: "Invoice No.:",
    B: invoice.invoiceNo,
  });
  invoiceTable.push({
    A: "Date Issued: ",
    B: `${invoice.createdDate} check!`,
  });
  invoiceTable.push({
    A: "Contract/Project Period:",
    B:
      `${moment(invoice.contractPeriod?.from).format("DD-MM-YYYY")} to ${moment(
        invoice.contractPeriod?.to
      ).format("DD-MM-YYYY")}` ||
      `${moment(invoice.projectPeriod?.from).format("DD-MM-YYYY")} to ${moment(
        invoice.projectPeriod?.to
      ).format("DD-MM-YYYY")}`,
  });

  const invoiceTableLines = invoiceTable.length;
  const spaceAfterInvoice = 2;

  let customerDetailTable = [];
  customerDetailTable.push({
    A: "Customer Details",
  });
  customerDetailTable.push({});
  customerDetailTable.push({
    A: "Customer Type:",
    B: customer.customerType,
  });
  customerDetailTable.push({
    A: "Company Name:",
    B: customer.company,
  });
  customerDetailTable.push({
    A: "UEN No.:",
    B: customer.uenNo,
  });
  customerDetailTable.push({
    A: "PIC Name:",
    B: customer.picName,
  });
  customerDetailTable.push({
    A: "PIC Email:",
    B: customer.picEmail,
  });
  customerDetailTable.push({
    A: "PIC Contact No.:",
    B: `(${customer.picContactNo.ext}) ${customer.picContactNo.phoneNumber}`,
  });
  customerDetailTable.push({
    A: "POC Name:",
    B: customer.pocName,
  });
  customerDetailTable.push({
    A: "POC Email:",
    B: customer.pocEmail,
  });
  customerDetailTable.push({
    A: "POC Contact No.:",
    B: `(${customer.pocContactNo.ext}) ${customer.pocContactNo.phoneNumber}`,
  });

  const customerDetailTableLines = customerDetailTable.length;
  const spaceAfterCustomerDetail = 2;

  // **
  // Billing Address
  // **

  let billingAddressTable = [];
  billingAddressTable.push({
    A: "Billing Address",
  });
  billingAddressTable.push({});
  billingAddressTable.push({
    A: "Address:",
    B: billingAddress.addressLine1,
    C: billingAddress.addressLine2,
  });
  billingAddressTable.push({
    A: "Postal Code:",
    B: billingAddress.postalCode,
  });

  const billingAddressTableLines = billingAddressTable.length;
  const spaceAfterBillingAddress = 4;

  // **
  // Project Detail
  // **

  let projectDetailTable = [];
  projectDetailTable.push({
    A: "Project Details",
  });

  const projectDetailTableLines = projectDetailTable.length;
  const spaceAfterProjectDetail = 1;

  // **
  // Project Location
  // **

  let projectLocationTable = [];
  let subProjectLocationTableLines = {};
  projectLocation.forEach((location, index) => {
    let subTable = [];
    subTable.push({
      A: `Project Location ${index + 1}`,
    });
    subTable.push({
      A: "Address:",
      B: location.addressLine1,
      C: location.addressLine2,
    });
    subTable.push({
      A: "Postal code:",
      B: location.postalCode,
    });
    subTable.push({});

    subProjectLocationTableLines = subTable.length;
    projectLocationTable = [...projectLocationTable, ...subTable];
  });

  const projectLocationTableLines = projectLocationTable.length;
  const spaceAfterProjectLocation = 1;

  // **
  // Door Type
  // **

  let doorTypeTable = [];
  let subDoorTypeTableLines;
  doorTypes.forEach((doorType) => {
    let subTable = [];
    subTable.push({
      A: `Door Type ${doorType?.door?.doorTypeIndex}`,
    });
    subTable.push({
      A: "Door Type:",
      B: doorType?.door?.doorType,
      C: doorType?.door?.doorTypeOtherText,
    });
    subTable.push({
      A: "Material of Door:",
      B: doorType?.door?.doorMaterial,
      C: doorType?.door?.doorMaterialOtherText,
    });
    subTable.push({
      A: "Type of Floor Guide:",
      B: doorType?.door?.floorGuideType,
      C: doorType?.door?.floorGuideTypeOtherText,
    });
    subTable.push({
      A: "Door Length:",
      B: doorType?.door?.doorSize?.length,
    });
    subTable.push({
      A: "Door Height:",
      B: doorType?.door?.doorSize?.height,
    });
    subTable.push({
      A: "Door Thickness:",
      B: doorType?.door?.doorSize?.thickness,
    });
    subTable.push({
      A: "Track Length:",
      B: doorType?.door?.doorSize?.trackLength,
    });
    subTable.push({
      A: "Cover Length:",
      B: doorType?.door?.doorSize?.coverLength,
    });
    subTable.push({
      A: "Glass Clamp Size:",
      B: doorType?.door?.doorSize?.glassClampSize,
    });
    subTable.push({});

    subDoorTypeTableLines = subTable.length;
    doorTypeTable = [...doorTypeTable, ...subTable];
  });

  const doorTypeTableLines = doorTypeTable.length;
  const spaceAfterDoorType = 1;

  // **
  // Inventories
  // **

  let inventoryTable = [];
  let subInventoryTableLines;
  inventories.forEach((inventory, index) => {
    let subTable = [];
    subTable.push({
      A: `Door Type ${index + 1} Inventory List`,
    });
    subTable.push({});
    subTable.push({
      A: "Inventory Name",
      B: "Category",
      C: "Location",
      D: "SKU",
      E: "Qty",
    });
    subTable.push({
      A: inventory?.inventoryName,
      B: inventory?.categoryName,
      C: inventory?.locationName,
      D: inventory?.sku,
      E: inventory?.qty,
    });
    subTable.push({});

    subInventoryTableLines = subTable.length;
    inventoryTable = [...inventoryTable, ...subTable];
  });

  const inventoryTableLines = inventoryTable.length;
  const spaceAfterInventory = 1;

  // **
  // Products
  // **

  let productTable = [];
  let totalLinesBeforeProductSumary;
  let subTotal = 0;
  // count start at product service/not row 1
  productTable.push({
    A: "Products/Services",
  });
  productTable.push({});
  productTable.push({
    A: "Product/Service",
    B: "Qty",
    C: "Unit",
    D: "Price",
  });
  products.forEach((product) => {
    productTable.push({
      A: product.productName,
      B: product.qty,
      C: product.unit,
      D: product.price,
    });
    subTotal += product.qty * product.price;
  });
  totalLinesBeforeProductSumary = productTable.length;
  productTable.push({
    C: "Subtotal",
    D: `$${subTotal}`,
  });
  const gstValue = (subTotal * invoice.gst) / 100;
  productTable.push({
    C: `GST ${invoice.gst}%`,
    D: `$${gstValue}`,
  });
  productTable.push({
    C: "Discount",
    D: `$${invoice.discount}`,
  });
  productTable.push({
    C: "Total",
    D: `$${subTotal + gstValue - invoice.discount}`,
  });

  const productTableLines = productTable.length;
  const spaceAfterProduct = 2;

  // **
  // Remarks
  // **

  let remarkTable = [];
  remarkTable.push({
    A: "Remarks",
  });
  remarkTable.push({
    A: remarks,
  });

  const remarkTableLines = remarkTable.length;
  const spaceAfterRemark = 2;

  // **
  // Terms
  // **

  let termTable = [];
  termTable.push({
    A: "Terms",
  });
  termTable.push({
    A: terms,
  });

  const termTableLines = remarkTable.length;
  const spaceAfterTerm = 2;

  // **
  // Warranty
  // **

  let warrantyTable = [];
  warrantyTable.push({
    A: "Warranty",
  });
  warrantyTable.push({});
  warrantyTable.push({ A: "PIC Name:", B: warranty.pic });
  warrantyTable.push({ A: "PIC Email:", B: warranty.picEmail });
  warrantyTable.push({ A: "Location:", B: warranty.location });
  warrantyTable.push({
    A: "Date:",
    B: `${moment(warranty.startDate).format("DD-MM-YYYY")} to ${moment(
      warranty.endDate
    ).format("DD-MM-YYYY")}`,
  });

  const warrantyTableLines = warrantyTable.length;
  const spaceAfterWararanty = 2;

  // **
  // Payment term
  // **

  let paymentTermTable = [];
  paymentTermTable.push({
    A: "Payment Terms",
  });
  paymentTermTable.push({
    A: invoicePaymentTerm,
  });
  paymentTermTable.push({});

  const paymentTermTableLines = paymentTermTable.length;
  const spaceAfterPaymentTerm = 4;

  // **
  // Sub invoice detail
  // **

  let subInvoiceTable = [];
  subInvoiceTable.push({
    A: "Sub Invoice",
  });
  subInvoiceTable.push({});
  subInvoiceTable.push({
    A: "Sub Invoice No.:",
    B: subInvoice.invoiceNo,
  });
  subInvoiceTable.push({
    A: "Date Issued: ",
    B: `${moment(subInvoice.dateIssued).format("DD-MM-YYYY")}`,
  });

  const subInvoiceTableLines = subInvoiceTable.length;
  const spaceAfterSubInvoice = 2;

  // **
  // Sub invoice products
  // **

  let subInvoiceProductTable = [];
  let subInvoiceProductTotal = 0;
  let totalLinesBeforeSubInvoiceProductSumary;
  subInvoiceProductTable.push({
    A: "Products/Services",
  });
  subInvoiceProductTable.push({});
  subInvoiceProductTable.push({
    A: "Product/Service",
    B: "Qty",
    C: "Unit",
    D: "Price",
  });
  subInvoiceProductServices.forEach((product) => {
    subInvoiceProductTable.push({
      A: product.productName,
      B: product.qty,
      C: product.unit,
      D: product.price,
    });
    subInvoiceProductTotal += product.qty * product.price;
  });
  totalLinesBeforeSubInvoiceProductSumary = subInvoiceProductTable.length;
  subInvoiceProductTable.push({
    C: "Total",
    D: `$${subInvoiceProductTotal}`,
  });

  const subInvoiceProductTableLines = subInvoiceProductTable.length;
  const spaceAftersubInvoiceProduct = 2;

  // **
  // Sub invoice payment term
  // **

  let subInvoicetermTable = [];
  subInvoicetermTable.push({
    A: "Payment Terms",
  });
  subInvoicetermTable.push({
    A: subInvoicePaymentTerm,
  });
  subInvoicetermTable.push({});

  const subInvoicePaymentTermTableLines = subInvoicetermTable.length;
  const spaceAftersubInvoicePaymentTerm = 2;

  // **
  // Signature
  // **
  // Defined in addStyles function

  const finalData = [
    ...renderLineSpace(space0),
    ...companyInforTable,
    ...renderLineSpace(spaceAfterCompanyInfor),
    ...invoiceTable,
    ...renderLineSpace(spaceAfterInvoice),
    ...customerDetailTable,
    ...renderLineSpace(spaceAfterCustomerDetail),
    ...billingAddressTable,
    ...renderLineSpace(spaceAfterBillingAddress),
    ...projectDetailTable,
    ...renderLineSpace(spaceAfterProjectDetail),
    ...projectLocationTable,
    ...renderLineSpace(spaceAfterProjectLocation),
    ...doorTypeTable,
    ...renderLineSpace(spaceAfterDoorType),
    ...inventoryTable,
    ...renderLineSpace(spaceAfterInventory),
    ...productTable,
    ...renderLineSpace(spaceAfterProduct),
    ...remarkTable,
    ...renderLineSpace(spaceAfterRemark),
    ...termTable,
    ...renderLineSpace(spaceAfterTerm),
    ...warrantyTable,
    ...renderLineSpace(spaceAfterWararanty),
    ...paymentTermTable,
    ...renderLineSpace(spaceAfterPaymentTerm),
    ...subInvoiceTable,
    ...renderLineSpace(spaceAfterSubInvoice),
    ...subInvoiceProductTable,
    ...renderLineSpace(spaceAftersubInvoiceProduct),
    ...subInvoicetermTable,
    ...renderLineSpace(spaceAftersubInvoicePaymentTerm),
  ];
  // const file = XLSX.readFile(fileName);

  const workSheet = XLSX.utils.json_to_sheet(finalData, { skipHeader: true });
  const workBook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workBook, workSheet, "Sheet 1");
  XLSX.writeFile(workBook, fileName);

  const totalLinesBeforeInvoiceTable =
    space0 + companyInforLines + spaceAfterCompanyInfor;

  const totalLinesBeforeCustomerDetailTable =
    totalLinesBeforeInvoiceTable + invoiceTableLines + spaceAfterCustomerDetail;

  const totalLinesBeforeBillingAddressTable =
    totalLinesBeforeCustomerDetailTable +
    customerDetailTableLines +
    spaceAfterCustomerDetail;

  const totalLinesBeforeProjectDetailTable =
    totalLinesBeforeBillingAddressTable +
    billingAddressTableLines +
    spaceAfterBillingAddress;

  const totalLinesBeforeProjectLocationTable =
    totalLinesBeforeProjectDetailTable +
    projectDetailTableLines +
    spaceAfterProjectDetail;

  const getProjectLocationHeaders = (projectLocation) => {
    let headers = {};
    projectLocation.forEach((_, index) => {
      headers[`projectLocationHeader${index}`] = `A${
        totalLinesBeforeProjectLocationTable +
        1 +
        subProjectLocationTableLines * index
      }:C${
        totalLinesBeforeProjectLocationTable +
        1 +
        subProjectLocationTableLines * index
      }`;
    });
    return headers;
  };

  const totalLinesBeforeDoorTypeTable =
    totalLinesBeforeProjectLocationTable +
    projectLocationTableLines +
    spaceAfterProjectLocation;

  const getDoorTypeHeaders = (workOrder) => {
    let headers = {};
    workOrder.forEach((_, index) => {
      headers[`workOrderHeader${index}`] = `A${
        totalLinesBeforeDoorTypeTable + 1 + subDoorTypeTableLines * index
      }:C${totalLinesBeforeDoorTypeTable + 1 + subDoorTypeTableLines * index}`;
    });
    return headers;
  };

  const totalLinesBeforeInventoryTable =
    totalLinesBeforeDoorTypeTable + doorTypeTableLines + spaceAfterDoorType;

  const getInventoryHeaders = (inventories) => {
    let headers = {};
    let subHeaders = {};
    inventories.forEach((_, index) => {
      headers[`inventoryHeader${index}`] = `A${
        totalLinesBeforeInventoryTable + 1 + subInventoryTableLines * index
      }:E${
        totalLinesBeforeInventoryTable + 1 + subInventoryTableLines * index
      }`;
      subHeaders[`subInventoryHeader${index}`] = `A${
        totalLinesBeforeInventoryTable + 3 + subInventoryTableLines * index
      }:E${
        totalLinesBeforeInventoryTable + 3 + subInventoryTableLines * index
      }`;
    });

    return { ...headers, ...subHeaders };
  };

  const totalLinesBeforeProductTable =
    totalLinesBeforeInventoryTable + inventoryTableLines + spaceAfterInventory;

  const totalLinesBeforeRemarkTable =
    totalLinesBeforeProductTable + productTableLines + spaceAfterProduct;

  const totalLinesBeforeTermTable =
    totalLinesBeforeRemarkTable + remarkTableLines + spaceAfterRemark;

  const totalLinesBeforeWarrantyTable =
    totalLinesBeforeTermTable + termTableLines + spaceAfterTerm;

  const totalLinesBeforePaymentTermTable =
    totalLinesBeforeWarrantyTable + warrantyTableLines + spaceAfterWararanty;

  const totalLinesBeforeSubInvoiceTable =
    totalLinesBeforePaymentTermTable +
    paymentTermTableLines +
    spaceAfterPaymentTerm;

  const totalLinesBeforeSubInvoiceProductTable =
    totalLinesBeforeSubInvoiceTable +
    subInvoiceTableLines +
    spaceAfterSubInvoice;

  const totalLinesBeforeSubInvoicePaymentTermTable =
    totalLinesBeforeSubInvoiceProductTable +
    subInvoiceProductTableLines +
    spaceAftersubInvoiceProduct;

  const totalLineBeforeSignature =
    totalLinesBeforeSubInvoicePaymentTermTable +
    subInvoicePaymentTermTableLines +
    spaceAftersubInvoicePaymentTerm;

  const dataInfo = {
    globalStyle: "A1:Z999",
    invoiceHeader: `A${totalLinesBeforeInvoiceTable + 1}:C${
      totalLinesBeforeInvoiceTable + 1
    }`,
    customerDetailHeader: `A${totalLinesBeforeCustomerDetailTable + 1}:C${
      totalLinesBeforeCustomerDetailTable + 1
    }`,
    billingAddressHeader: `A${totalLinesBeforeBillingAddressTable + 1}:C${
      totalLinesBeforeBillingAddressTable + 1
    }`,
    projectDetailHeader: `A${totalLinesBeforeProjectDetailTable + 1}:E${
      totalLinesBeforeProjectDetailTable + 1
    }`,
    ...getProjectLocationHeaders(projectLocation),
    ...getDoorTypeHeaders(doorTypes),
    ...getInventoryHeaders(inventories),
    productHeader: `A${totalLinesBeforeProductTable + 1}:E${
      totalLinesBeforeProductTable + 1
    }`,
    subProductHeader: `A${totalLinesBeforeProductTable + 3}:D${
      totalLinesBeforeProductTable + 3
    }`,
    sumaryProductHeader: `C${
      totalLinesBeforeProductTable + totalLinesBeforeProductSumary + 1
    }:C${totalLinesBeforeProductTable + totalLinesBeforeProductSumary + 4}`,
    totalOnProduct: `C${
      totalLinesBeforeProductTable + totalLinesBeforeProductSumary + 4
    }:D${totalLinesBeforeProductTable + totalLinesBeforeProductSumary + 4}`,
    totalLinesBeforeProductTable,
    totalLinesBeforeProductSumary,
    remarkHeader: `A${totalLinesBeforeRemarkTable + 1}:C${
      totalLinesBeforeRemarkTable + 1
    }`,
    remarkBody: `A${totalLinesBeforeRemarkTable + 2}:C${
      totalLinesBeforeRemarkTable + 2
    }`,
    totalLinesBeforeRemarkTable,
    termHeader: `A${totalLinesBeforeTermTable + 1}:C${
      totalLinesBeforeTermTable + 1
    }`,
    termBody: `A${totalLinesBeforeTermTable + 2}:C${
      totalLinesBeforeTermTable + 2
    }`,
    totalLinesBeforeTermTable,
    warrantyHeader: `A${totalLinesBeforeWarrantyTable + 1}:C${
      totalLinesBeforeWarrantyTable + 1
    }`,
    paymentTermHeader: `A${totalLinesBeforePaymentTermTable + 1}:C${
      totalLinesBeforePaymentTermTable + 1
    }`,
    paymentTermImg: `B${totalLinesBeforePaymentTermTable + 2}:C${
      totalLinesBeforePaymentTermTable + 3
    }`,
    subInvoiceHeader: `A${totalLinesBeforeSubInvoiceTable + 1}:C${
      totalLinesBeforeSubInvoiceTable + 1
    }`,
    subInvoiceProductHeader: `A${totalLinesBeforeSubInvoiceProductTable + 1}:C${
      totalLinesBeforeSubInvoiceProductTable + 1
    }`,
    subInvoiceSubProductHeader: `A${
      totalLinesBeforeSubInvoiceProductTable + 3
    }:D${totalLinesBeforeSubInvoiceProductTable + 3}`,
    totalSubInvoiceProduct: `C${
      totalLinesBeforeSubInvoiceProductTable +
      totalLinesBeforeSubInvoiceProductSumary +
      1
    }:D${
      totalLinesBeforeSubInvoiceProductTable +
      totalLinesBeforeSubInvoiceProductSumary +
      1
    }`,
    subInvoicePaymentTermHeader: `A${
      totalLinesBeforeSubInvoicePaymentTermTable + 1
    }:C${totalLinesBeforeSubInvoicePaymentTermTable + 1}`,
    subInvoicePaymentTermImg: `B${
      totalLinesBeforeSubInvoicePaymentTermTable + 2
    }:C${totalLinesBeforeSubInvoicePaymentTermTable + 3}`,
    totalLineBeforeSignature,
  };

  return addStyles(dataInfo);
};

const addStyles = (dataInfo) => {
  return XlsxPopulate.fromFileAsync(fileName).then((workbook) => {
    workbook.sheets().forEach((sheet) => {
      sheet.column("A").width(30);
      sheet.column("B").width(30);
      sheet.column("C").width(30);
      sheet.column("D").width(30);
      sheet.column("E").width(30);
      sheet.column("F").width(30);
      sheet.column("G").width(30);
      sheet.column("H").width(30);

      sheet.range(dataInfo.globalStyle).style({
        horizontalAlignment: "left",
      });
      sheet.range(dataInfo.invoiceHeader).merged(true).style({
        bold: true,
        fontSize: 24,
        bottomBorder: true,
        bottomBorderColor: "000000",
        bottomBorderStyle: "thin",
      });
      sheet.range(dataInfo.customerDetailHeader).merged(true).style({
        bold: true,
        fontSize: 24,
        bottomBorder: true,
        bottomBorderColor: "000000",
        bottomBorderStyle: "thin",
      });
      sheet.range(dataInfo.billingAddressHeader).merged(true).style({
        bold: true,
        fontSize: 24,
        bottomBorder: true,
        bottomBorderColor: "000000",
        bottomBorderStyle: "thin",
      });
      sheet.range(dataInfo.projectDetailHeader).merged(true).style({
        bold: true,
        fontSize: 24,
        bottomBorder: true,
        bottomBorderColor: "000000",
        bottomBorderStyle: "thin",
      });
      projectLocation.forEach((_, index) => {
        sheet
          .range(dataInfo[`projectLocationHeader${index}`])
          .merged(true)
          .style({
            bold: true,
            fontSize: 12,
            fill: "C8C8C8",
          });
      });
      doorTypes.forEach((_, index) => {
        sheet.range(dataInfo[`workOrderHeader${index}`]).merged(true).style({
          bold: true,
          fontSize: 12,
          fill: "C8C8C8",
        });
      });
      inventories.forEach((_, index) => {
        sheet.range(dataInfo[`inventoryHeader${index}`]).merged(true).style({
          bold: true,
          fontSize: 12,
          bottomBorder: true,
          bottomBorderColor: "000000",
          bottomBorderStyle: "thin",
        });
        sheet.range(dataInfo[`subInventoryHeader${index}`]).style({
          bold: true,
          fontSize: 12,
          fill: "C8C8C8",
        });
      });
      sheet.range(dataInfo.productHeader).merged(true).style({
        bold: true,
        fontSize: 12,
        bottomBorder: true,
        bottomBorderColor: "000000",
        bottomBorderStyle: "thin",
      });
      sheet.range(dataInfo.subProductHeader).style({
        bold: true,
        fontSize: 12,
        fill: "C8C8C8",
      });
      sheet.range(dataInfo.sumaryProductHeader).style({
        bold: true,
        fontSize: 12,
      });
      sheet.range(dataInfo.totalOnProduct).style({
        bold: true,
        fontSize: 12,
        fontColor: "FF0000",
        bottomBorder: true,
        bottomBorderColor: "000000",
        bottomBorderStyle: "double",
        topBorder: true,
        topBorderColor: "000000",
        topBorderStyle: "thin",
      });
      sheet.range(dataInfo.remarkHeader).merged(true).style({
        bold: true,
        fontSize: 14,
        fill: "C8C8C8",
      });
      sheet.range(dataInfo.remarkBody).merged(true).style({
        fontSize: 12,
        verticalAlignment: "top",
      });
      sheet.row(dataInfo.totalLinesBeforeRemarkTable + 2).height(70);
      sheet.range(dataInfo.termHeader).merged(true).style({
        bold: true,
        fontSize: 14,
        fill: "C8C8C8",
      });
      sheet.range(dataInfo.termBody).merged(true).style({
        fontSize: 12,
        verticalAlignment: "top",
      });
      sheet.row(dataInfo.totalLinesBeforeTermTable + 2).height(70);
      sheet.range(dataInfo.warrantyHeader).merged(true).style({
        bold: true,
        fontSize: 14,
        bottomBorder: true,
        bottomBorderColor: "000000",
        bottomBorderStyle: "thin",
      });
      sheet.range(dataInfo.paymentTermHeader).merged(true).style({
        bold: true,
        fontSize: 14,
        bottomBorder: true,
        bottomBorderColor: "000000",
        bottomBorderStyle: "thin",
      });
      sheet
        .range(dataInfo.paymentTermImg)
        .merged(true)
        .style({
          verticalAlignment: "center",
          horizontalAlignment: "center",
        })
        .value("(image)");

      sheet.range(dataInfo.subInvoiceHeader).merged(true).style({
        bold: true,
        fontSize: 18,
        bottomBorder: true,
        bottomBorderColor: "000000",
        bottomBorderStyle: "thin",
      });
      sheet.range(dataInfo.subInvoiceProductHeader).merged(true).style({
        bold: true,
        fontSize: 12,
        bottomBorder: true,
        bottomBorderColor: "000000",
        bottomBorderStyle: "thin",
      });
      sheet.range(dataInfo.subInvoiceSubProductHeader).style({
        bold: true,
        fontSize: 12,
        fill: "C8C8C8",
      });
      sheet.range(dataInfo.totalSubInvoiceProduct).style({
        bold: true,
        fontSize: 12,
        fontColor: "FF0000",
        bottomBorder: true,
        bottomBorderColor: "000000",
        bottomBorderStyle: "double",
        topBorder: true,
        topBorderColor: "000000",
        topBorderStyle: "thin",
      });
      sheet.range(dataInfo.subInvoicePaymentTermHeader).merged(true).style({
        bold: true,
        fontSize: 14,
        bottomBorder: true,
        bottomBorderColor: "000000",
        bottomBorderStyle: "thin",
      });
      sheet
        .range(dataInfo.subInvoicePaymentTermImg)
        .merged(true)
        .style({
          verticalAlignment: "center",
          horizontalAlignment: "center",
        })
        .value("(image)");

      // Signature
      sheet
        .range(
          `A${dataInfo.totalLineBeforeSignature + 1}:B${
            dataInfo.totalLineBeforeSignature + 1
          }`
        )
        .merged(true)
        .style({
          bold: true,
        })
        .value("Door Suisse Technology Pte Ltd");
      sheet
        .range(
          `D${dataInfo.totalLineBeforeSignature + 1}:E${
            dataInfo.totalLineBeforeSignature + 1
          }`
        )
        .merged(true)
        .style({
          bold: true,
        })
        .value("Accepted By");
      sheet.row(dataInfo.totalLineBeforeSignature + 2).height(70);
      sheet
        .range(
          `A${dataInfo.totalLineBeforeSignature + 2}:B${
            dataInfo.totalLineBeforeSignature + 2
          }`
        )
        .merged(true)
        .style({
          verticalAlignment: "center",
          horizontalAlignment: "center",
          bottomBorder: true,
          bottomBorderColor: "000000",
          bottomBorderStyle: "thin",
        })
        .value("Signature");
      sheet
        .range(
          `D${dataInfo.totalLineBeforeSignature + 2}:E${
            dataInfo.totalLineBeforeSignature + 2
          }`
        )
        .merged(true)
        .style({
          verticalAlignment: "center",
          horizontalAlignment: "center",
          bottomBorder: true,
          bottomBorderColor: "000000",
          bottomBorderStyle: "thin",
        })
        .value("Signature");
      sheet
        .range(
          `D${dataInfo.totalLineBeforeSignature + 3}:E${
            dataInfo.totalLineBeforeSignature + 3
          }`
        )
        .merged(true)
        .style({
          bold: true,
        })
        .value("Sign, Name, and Co. Stam");
    });
    return workbook.toFileAsync(fileName);
  });
};

const generateCSV = (fn) => {
  fs.unlink(fileName, (err) => {
    if (err) {
      fn({
        companyInfo,
        invoice,
        customer,
        billingAddress,
        projectLocation,
        doorTypes,
        inventories,
        products,
        remarks,
        terms,
        warranty,
        invoicePaymentTerm,
        subInvoice,
        subInvoiceProductServices,
        subInvoicePaymentTerm,
      });
    } else {
      fn({
        companyInfo,
        invoice,
        customer,
        billingAddress,
        projectLocation,
        doorTypes,
        inventories,
        products,
        remarks,
        terms,
        warranty,
        invoicePaymentTerm,
        subInvoice,
        subInvoiceProductServices,
        subInvoicePaymentTerm,
      });
    }
  });
};

generateCSV(handleExport);
