const XlsxPopulate = require("xlsx-populate");
const XLSX = require("xlsx");
const moment = require("moment");
const fs = require("fs");

// project belong

const projectNo = "PM3997";

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

// Work Order
const workOrders = [
  {
    doorTypeId: "0191afdb-22af-4313-833d-624d3d8c2458",
    workOrderNo: "WO7742",
    doorId: "8d9f6311-2132-4025-9cff-f7824713a4c4",
    doorNo: "PM3997-1",
    estCompletedDate: "2023-06-05T17:00:00.000Z",
    door: {
      doorTypeIndex: 1,
      doorType: "Single (E-STA)",
      doorMaterial: "Steel",
      floorGuideType: "FG-10",
      doorTypeOtherText: "",
      doorMaterialOtherText: "",
      floorGuideTypeOtherText: "",
      doorSize: {
        height: 1100,
        length: 800,
        thickness: 100,
        trackLength: 1200,
        coverLength: 1000,
        glassClampSize: 100,
      },
      remarks: "adfasd",
    },
    mechanism: [
      {
        inventoryId: "e6fdbb7-846e-4a94-9379-37eb81e0c4c8",
        categoryName: null,
        locationName: null,
      },
      {
        inventoryId: "e6fdbb7-846e-4a94-9379-37eb81e0c4c8",
        categoryName: null,
        locationName: null,
      },
    ],
    actuatingDevices: [
      {
        inventoryId: "e6fdbb7-846e-4a94-9379-37eb81e0c4c8",
        categoryName: null,
        locationName: null,
      },
    ],
    operatorCutHalf: true,
    coverCutHalf: false,
    completeWithEndCap: true,
    sectionDetail: "asdfasfasf",
    mountingDetail: {
      mounting: ["Transom"],
      otherText: "",
    },
    installation: {
      track: "",
      kitAccessories: "",
      doorLeaf: "",
      otherReturnTrip: "",
    },
    comment: "afasdfasdfsf",
    doorType: {
      isActive: true,
      isDeleted: false,
      _id: "648fcf5d03cf3e4d836ea3e7",
      id: "0191afdb-22af-4313-833d-624d3d8c2458",
      createdDate: "2023-06-19T03:45:33.258Z",
      updatedDate: "2023-06-19T03:45:33.258Z",
      count: 1,
      doorType: "Single (E-STA)",
      doorTypeOtherText: "",
      doorMaterial: "Steel",
      doorMaterialOtherText: "",
      floorGuideType: "FG-10",
      floorGuideTypeOtherText: "",
      remarks: "adfasd",
      doorSize: {
        height: 1100,
        length: 800,
        thickness: 100,
        trackLength: 1200,
        coverLength: 1000,
        glassClampSize: 100,
      },
    },
  },
  {
    doorTypeId: "94614a19-33a6-4820-89f2-344c2129514d",
    workOrderNo: "WO2816",
    doorId: "8f8ab618-1c4c-46c8-97fc-8b6c2743f0a3",
    doorNo: "PM3997-2",
    estCompletedDate: "2023-06-04T17:00:00.000Z",
    door: {
      doorTypeIndex: 1,
      doorType: "Telescopic (D-TSA)",
      doorMaterial: "Steel",
      floorGuideType: "FG-10",
      doorTypeOtherText: "",
      doorMaterialOtherText: "",
      floorGuideTypeOtherText: "",
      doorSize: {
        height: 1000,
        length: 800,
        thickness: 100,
        trackLength: 1100,
        coverLength: 1100,
        glassClampSize: 120,
      },
      remarks: "",
    },
    mechanism: [
      {
        inventoryId: "e6fdbb7-846e-4a94-9379-37eb81e0c4c8",
        categoryName: null,
        locationName: null,
      },
      {
        inventoryId: "e6fdbb7-846e-4a94-9379-37eb81e0c4c8",
        categoryName: null,
        locationName: null,
      },
    ],
    actuatingDevices: [
      {
        inventoryId: "e6fdbb7-846e-4a94-9379-37eb81e0c4c8",
        categoryName: null,
        locationName: null,
      },
    ],
    operatorCutHalf: true,
    coverCutHalf: false,
    completeWithEndCap: false,
    sectionDetail: "adfasfasdfasf",
    mountingDetail: {
      mounting: ["Conceal"],
      otherText: "",
    },
    installation: {
      track: "sdfs",
      kitAccessories: "",
      doorLeaf: "",
      otherReturnTrip: "",
    },
    comment: "adsfasdfasdfas",
    doorType: {
      isActive: true,
      isDeleted: false,
      _id: "648fcf5d03cf3e4d836ea3e8",
      id: "94614a19-33a6-4820-89f2-344c2129514d",
      createdDate: "2023-06-19T03:45:33.259Z",
      updatedDate: "2023-06-19T03:45:33.259Z",
      count: 1,
      doorType: "Telescopic (D-TSA)",
      doorTypeOtherText: "",
      doorMaterial: "Steel",
      doorMaterialOtherText: "",
      floorGuideType: "FG-10",
      floorGuideTypeOtherText: "",
      remarks: "",
      doorSize: {
        height: 1000,
        length: 800,
        thickness: 100,
        trackLength: 1100,
        coverLength: 1100,
        glassClampSize: 120,
      },
    },
  },
];

const renderLineSpace = (lines) => {
  let emptyArray = [];
  for (let i = 0; i < lines; i++) {
    emptyArray.push("");
  }
  return emptyArray;
};

const fileName = `workOrder-${projectNo}.xlsx`;

const handleExport = ({ companyInfo }) => {
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

  let workOrderTables = [];
  let subWorkOrderProperties = [];
  workOrders.forEach((workOrder) => {
    let workOrderTable = [];
    workOrderTable.push({
      A: `Works Order Form No.: ${workOrder.workOrderNo}`,
    });
    let spaceAfterWorkOrderHeader = 1;
    // **
    // Estimate completed date
    // **
    let estCompleteDateTable = [];
    estCompleteDateTable.push({
      A: "To be completed by:",
      B: `${moment(workOrder.estCompletedDate).format("DD-MM-YYYY")}`,
    });
    let spaceAfterEstCompleteDate = 1;

    // **
    // Door detail
    // **
    let doorDetailTable = [];
    doorDetailTable.push({
      A: "Door No.:",
      B: workOrder?.doorNo,
    });
    doorDetailTable.push({
      A: "Door type:",
      B: workOrder?.door?.doorType,
    });
    doorDetailTable.push({
      A: "Material of Door:",
      B: workOrder?.door?.doorMaterial,
    });
    doorDetailTable.push({
      A: "Type of Floor Guide:",
      B: workOrder?.door?.floorGuideType,
    });
    doorDetailTable.push({
      A: "Door Length:",
      B: workOrder?.door?.doorSize?.length,
    });
    doorDetailTable.push({
      A: "Door Height:",
      B: workOrder?.door?.doorSize?.height,
    });
    doorDetailTable.push({
      A: "Door Thickness:",
      B: workOrder?.door?.doorSize?.thickness,
    });
    doorDetailTable.push({
      A: "Track Length:",
      B: workOrder?.door?.doorSize?.trackLength,
      C: "Operator to cut half:",
      D: workOrder?.door?.operatorCutHalf ? "Yes" : "No",
    });
    doorDetailTable.push({
      A: "Cover Length:",
      B: workOrder?.door?.doorSize?.coverLength,
      C: "Cover to cut half:",
      D: workOrder?.door?.coverCutHalf ? "Yes" : "No",
    });
    doorDetailTable.push({
      C: "Complete with end Caps",
      D: workOrder?.door?.completeWithEndCap ? "Yes" : "No",
    });
    doorDetailTable.push({
      A: "Glass Clamp Size:",
      B: workOrder?.door?.doorSize?.glassClampSize,
    });

    let spaceAfterDoorDetail = 2;

    // **
    // Section detail
    // **

    let sectionDetailTable = [];
    sectionDetailTable.push({
      A: "Section Detail:",
    });
    sectionDetailTable.push({
      A: workOrder?.sectionDetail,
    });

    let spaceAfterSectionDetail = 2;

    // **
    // Mounting details
    // **

    let mountingDetailTable = [];
    mountingDetailTable.push({
      A: "Mounting Details",
    });
    mountingDetailTable.push({
      A: "Mounting:",
      B: workOrder?.mountingDetail?.mounting.join(", "),
    });
    mountingDetailTable.push({
      A: "Other:",
      B: workOrder?.mountingDetail?.otherText,
    });

    let spaceAfterMountingDetail = 2;

    // **
    // Installation Information
    // **

    let installationTable = [];
    installationTable.push({
      A: "Installation Information",
    });
    Object.entries(workOrder.installation).forEach(([key, value]) => {
      installationTable.push({
        A: key,
        B: value,
      });
    });

    let spaceAfterInstallation = 2;

    // **
    // Comment
    // **

    let commentTable = [];
    commentTable.push({
      A: "Comment",
    });
    commentTable.push({
      A: workOrder?.sectionDetail,
    });

    let spaceAfterComment = 2;

    workOrderTable = [
      ...workOrderTable,
      ...renderLineSpace(spaceAfterWorkOrderHeader),
      ...estCompleteDateTable,
      ...renderLineSpace(spaceAfterEstCompleteDate),
      ...doorDetailTable,
      ...renderLineSpace(spaceAfterDoorDetail),
      ...sectionDetailTable,
      ...renderLineSpace(spaceAfterSectionDetail),
      ...mountingDetailTable,
      ...renderLineSpace(spaceAfterMountingDetail),
      ...installationTable,
      ...renderLineSpace(spaceAfterInstallation),
      ...commentTable,
      ...renderLineSpace(spaceAfterComment),
    ];

    const totalLinesBeforeEstCompleteDate = spaceAfterWorkOrderHeader;
    const totalLinesBeforeDoorDetail =
      totalLinesBeforeEstCompleteDate +
      estCompleteDateTable.length +
      spaceAfterEstCompleteDate;
    const totalLinesBeforeSectionDetail =
      totalLinesBeforeDoorDetail +
      doorDetailTable.length +
      spaceAfterDoorDetail;
    const totalLinesBeforeMountingDetail =
      totalLinesBeforeSectionDetail +
      sectionDetailTable.length +
      spaceAfterSectionDetail;

    const totalLinesBeforeInstallation =
      totalLinesBeforeMountingDetail +
      mountingDetailTable.length +
      spaceAfterMountingDetail;

    const totalLinesBeforeComment =
      totalLinesBeforeInstallation +
      installationTable.length +
      spaceAfterInstallation;

    subWorkOrderProperties.push({
      length: workOrderTable.length,
      estComplete: { before: totalLinesBeforeEstCompleteDate },
      doorDetail: {
        before: totalLinesBeforeDoorDetail,
        length: doorDetailTable.length,
      },
      sectionDetail: {
        before: totalLinesBeforeSectionDetail,
      },
      mountingDetail: {
        before: totalLinesBeforeMountingDetail,
      },
      installation: {
        before: totalLinesBeforeInstallation,
      },
      comment: {
        before: totalLinesBeforeComment,
      },
    });
    workOrderTables = [...workOrderTables, ...workOrderTable];
  });

  const finalData = [
    ...renderLineSpace(space0),
    ...companyInforTable,
    ...renderLineSpace(spaceAfterCompanyInfor),
    ...workOrderTables,
  ];

  const workSheet = XLSX.utils.json_to_sheet(finalData, { skipHeader: true });
  const workBook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workBook, workSheet, "Sheet 1");
  XLSX.writeFile(workBook, fileName);

  const totalLinesBeforeWorkOrderTables =
    space0 + companyInforLines + spaceAfterCompanyInfor;

  const workOrderHeaders = (subWorkOrderProperties) => {
    let headers = {};
    let totalLength = 0;
    subWorkOrderProperties.forEach((property, index) => {
      const { length } = property;
      headers[`workOrder${index}`] = `A${
        totalLinesBeforeWorkOrderTables + 1 + totalLength
      }:C${totalLinesBeforeWorkOrderTables + 1 + totalLength}`;
      totalLength += length;
    });
    return headers;
  };

  const estCompleteDateHeaders = (subWorkOrderProperties) => {
    let headers = {};
    let totalLength = 0;
    subWorkOrderProperties.forEach((property, index) => {
      const { estComplete, length } = property;
      const totalLinesBeforeEstHeader =
        totalLinesBeforeWorkOrderTables + 1 + estComplete.before;
      headers[`estComplete${index}`] = `A${
        totalLinesBeforeEstHeader + 1 + totalLength
      }:A${totalLinesBeforeEstHeader + 1 + totalLength}`;
      totalLength += length;
    });
    return headers;
  };

  const doorDetailHeaders = (subWorkOrderProperties) => {
    let headers = {};
    let totalLength = 0;
    subWorkOrderProperties.forEach((property, index) => {
      const { doorDetail, length } = property;
      const totalLinesBeforeDoorDetail =
        totalLinesBeforeWorkOrderTables + 1 + doorDetail.before;
      headers[`doorDetail${index}`] = `A${
        totalLinesBeforeDoorDetail + 1 + totalLength
      }:A${totalLinesBeforeDoorDetail + 1 + totalLength + doorDetail.length}`;
      headers[`doorDetailCheck${index}`] = `C${
        totalLinesBeforeDoorDetail + 1 + totalLength
      }:C${totalLinesBeforeDoorDetail + 1 + totalLength + doorDetail.length}`;
      totalLength += length;
    });
    return headers;
  };

  const sectionDetailHeaders = (subWorkOrderProperties) => {
    let headers = {};
    let totalLength = 0;
    subWorkOrderProperties.forEach((property, index) => {
      const { sectionDetail, length } = property;
      const totalLinesBeforeSectionDetail =
        totalLinesBeforeWorkOrderTables + 1 + sectionDetail.before;
      headers[`sectionDetail${index}`] = `A${
        totalLinesBeforeSectionDetail + 1 + totalLength
      }:C${totalLinesBeforeSectionDetail + 1 + totalLength}`;
      headers[`sectionDetailContent${index}`] = `A${
        totalLinesBeforeSectionDetail + 2 + totalLength
      }:C${totalLinesBeforeSectionDetail + 2 + totalLength}`;
      totalLength += length;
    });
    return headers;
  };

  const getRowForSectionContent = (subWorkOrderProperties) => {
    let totalLength = 0;
    let rows = [];
    subWorkOrderProperties.forEach((property, index) => {
      const { sectionDetail, length } = property;
      const totalLinesBeforeSectionDetail =
        totalLinesBeforeWorkOrderTables + 1 + sectionDetail.before;
      rows.push(totalLinesBeforeSectionDetail + 2 + totalLength);
      totalLength += length;
    });
    return rows;
  };

  const mountingDetailHeaders = (subWorkOrderProperties) => {
    let headers = {};
    let totalLength = 0;
    subWorkOrderProperties.forEach((property, index) => {
      const { mountingDetail, length } = property;
      const totalLinesBeforeMountingDetailTables =
        totalLinesBeforeWorkOrderTables + 1 + mountingDetail.before;
      headers[`mountingDetail${index}`] = `A${
        totalLinesBeforeMountingDetailTables + 1 + totalLength
      }:C${totalLinesBeforeMountingDetailTables + 1 + totalLength}`;
      totalLength += length;
    });
    return headers;
  };

  const installationlHeaders = (subWorkOrderProperties) => {
    let headers = {};
    let totalLength = 0;
    subWorkOrderProperties.forEach((property, index) => {
      const { installation, length } = property;
      const totalLinesBeforeInstallationTables =
        totalLinesBeforeWorkOrderTables + 1 + installation.before;
      headers[`installation${index}`] = `A${
        totalLinesBeforeInstallationTables + 1 + totalLength
      }:C${totalLinesBeforeInstallationTables + 1 + totalLength}`;
      totalLength += length;
    });
    return headers;
  };

  const commentHeaders = (subWorkOrderProperties) => {
    let headers = {};
    let totalLength = 0;
    subWorkOrderProperties.forEach((property, index) => {
      const { comment, length } = property;
      const totalLinesBeforeComment =
        totalLinesBeforeWorkOrderTables + 1 + comment.before;
      headers[`comment${index}`] = `A${
        totalLinesBeforeComment + 1 + totalLength
      }:C${totalLinesBeforeComment + 1 + totalLength}`;
      headers[`commentContent${index}`] = `A${
        totalLinesBeforeComment + 2 + totalLength
      }:C${totalLinesBeforeComment + 2 + totalLength}`;
      totalLength += length;
    });
    return headers;
  };

  const getRowForCommentContent = (subWorkOrderProperties) => {
    let totalLength = 0;
    let rows = [];
    subWorkOrderProperties.forEach((property, index) => {
      const { comment, length } = property;
      const totalLinesBeforeComment =
        totalLinesBeforeWorkOrderTables + 1 + comment.before;
      rows.push(totalLinesBeforeComment + 2 + totalLength);
      totalLength += length;
    });
    return rows;
  };

  const dataInfo = {
    globalStyle: "A1:Z999",
    ...workOrderHeaders(subWorkOrderProperties),
    ...estCompleteDateHeaders(subWorkOrderProperties),
    ...doorDetailHeaders(subWorkOrderProperties),
    ...sectionDetailHeaders(subWorkOrderProperties),
    rowForSectionContents: getRowForSectionContent(subWorkOrderProperties),
    ...mountingDetailHeaders(subWorkOrderProperties),
    ...installationlHeaders(subWorkOrderProperties),
    ...commentHeaders(subWorkOrderProperties),
    rowForCommentContents: getRowForCommentContent(subWorkOrderProperties),
  };
  console.log(dataInfo);

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
      workOrders.forEach((_, index) => {
        sheet.range(dataInfo[`workOrder${index}`]).merged(true).style({
          bold: true,
          fontSize: 18,
          bottomBorder: true,
          bottomBorderColor: "000000",
          bottomBorderStyle: "thin",
        });
        sheet.range(dataInfo[`estComplete${index}`]).style({
          bold: true,
          fontSize: 12,
        });
        sheet.range(dataInfo[`doorDetail${index}`]).style({
          bold: true,
          fontSize: 12,
        });
        sheet.range(dataInfo[`doorDetailCheck${index}`]).style({
          bold: true,
          fontSize: 12,
        });
        sheet.range(dataInfo[`sectionDetail${index}`]).merged(true).style({
          bold: true,
          fontSize: 12,
          fill: "C8C8C8",
        });
        sheet
          .range(dataInfo[`sectionDetailContent${index}`])
          .merged(true)
          .style({
            verticalAlignment: "top",
            horizontalAlignment: "left",
          });
        sheet.range(dataInfo[`mountingDetail${index}`]).merged(true).style({
          bold: true,
          fontSize: 12,
          fill: "C8C8C8",
        });
        sheet.range(dataInfo[`installation${index}`]).merged(true).style({
          bold: true,
          fontSize: 12,
          fill: "C8C8C8",
        });
        sheet.range(dataInfo[`comment${index}`]).merged(true).style({
          bold: true,
          fontSize: 12,
          fill: "C8C8C8",
        });
        sheet.range(dataInfo[`commentContent${index}`]).merged(true).style({
          verticalAlignment: "top",
          horizontalAlignment: "left",
        });
      });

      dataInfo.rowForSectionContents.forEach((row) => {
        sheet.row(row).height(70);
      });
      dataInfo.rowForCommentContents.forEach((row) => {
        sheet.row(row).height(70);
      });
    });
    return workbook.toFileAsync(fileName);
  });
};

const generateCSV = (fn) => {
  fs.unlink(fileName, (err) => {
    if (err) {
      fn({
        companyInfo,
      });
    } else {
      fn({
        companyInfo,
      });
    }
  });
};

generateCSV(handleExport);
