import { getDateRange } from './date'

const defaultParams = {
  paramBlock: {
    data: {
      purchaseTerminationStatus: "0",
      isPlanNewestStatusQuery: "1",
      status: null,
      orderApproveTimeSection: null,
      itemAddr: "",
      isLogisticsPurchaseOrg: "1",
      ids: null,
      hideFileNames: "materialDesc,fbk72,fbk73,fbk77,fbk76,fbk75,fbk74",
      exportFields: "requestPlanNumber,fbk10,requestPlanName,planSourcesType,status,businessTypes,supplementaryRecordForBidding,fbk71,planQuantity,companyStock,companyPrice,planDeliveryDate,planVoucherPeople,planVoucherDate,planApproveTime,overseasPlanNumber,planIsEmergency,fbk113,purchaseTerminationStatus,purchaseTerminationPerson,purchaseTerminationTime,terminationReason,company,fbk28,materialNumber,fbk110,remarkSupplier,remarkPurchase,categoryCode,categoryName,measuringUnit,submissionName,submissionDepartment,fbk35,dlsbgs,dlsbbm,isOutlands,itemAddr,itemNumber,itemName,isAirTransport,fbk59,purchaseOrg,purchaseOrgName,productUse,planIsDelete,planPoolCreateTime,manuallyAssignPurchaseOrgName,manuallyAssignPurchaseOrgTime,manuallyAssignPurchaseGroupName,firstManuallyAssignPurchaseGroupTime,manuallyAssignPurchaseGroupTime,manuallyAssignPurchaserName,manuallyAssignPurchaserTime,fbk54,fbk55,fbk53,planNumber,fbk109,purchaseMode,fbk56,fbk58,fbk70,programmeIsDelete,programmeVoucherPeople,programmeVoucherDate,fbk46,fbk30,programmeApproveTime,fbk51,fbk52,enquiryNumber,enquiryIsDeleted,enquiryCreator,enquiryCreateTime,fbk40,fbk11,fbk14,enquiryApproveTime,supplierCode,supplierName,biddingNumber,realityOpenBidTime,fbk81,fbk27,bidApproveTime,bidIsDeleted,fbk25,fbk44,fbk43,isFrame,purchaseGroup,orderNumber,orderMaterialNumber,orderMaterialDesc,fbk26,factory,sapPurchaseNumber,orderIsDeleted,isManullyOrder,isContractOrder,fbk48,supplyQuantity,customsType,customsName,totalTaxAmount,currency,fbk95,deliverTime,fbk111,paymentWay,fbk61,isWeigh,isSampling,isMsds,sellSupplierCode,sellCompanyName,purchaseCompany,fbk23,sellPhone,purchaseHold,purchaseDeliver,orderMake,orderMakeDate,fbk45,orderApproveTime,scanUploadTime,scanModifyTime,relationOrderNumber,relationSapPurchaseNumber,relationOrderAccount,relationOrderMakeDate,relationOrderApproveTime,relationOrderIsDelete,synContractTime,contractNumber,contractSignTime,frameContractNumber,fbk112,deliveryTime,confirmArriveDate,deliveryNumber,deliveryCreator,deliveryQuantity,voucherNumber,voucherDate,receiveQuantity,unreceiveQuantity,relationVoucherNumber,relationVoucherDate,relationReceiveQuantity,relationUnreceiveQuantity,leConfirmArriveDate,bgOrderNumber,ducId,departTime,arrivalTime,receiptTime,portTime,approachTime,invoiceNo,invoiceRegisterTime,invoicePaidAmount,invoiceUnpaidAmount,emConsumeQuote,emAttachment,emProductType,emIsInventoryMaterial,emOrderSn,emDetailId,emTaxPrice,emTaxRate,emTotalTaxPrice,emNetPrice,emNetTotal,emOrderDetailOfferId,emMaterialName,fbk57,isUnusual,unusualRemark,unusualCreator,registrantTime,isImportant,importantRemark,importantRegistrantName,importantRegistrantTime,fbk67,fbk68,fbk69,fbk80,fbk83,fbk84,fbk85,fbk86,fbk89,fbk90,fbk91,fbk92,fbk93,fbk94,isPurchaseDone,fbk96,fbk97,fbk98,fbk105,fbk101,fbk106,fbk107,fbk108,sapPRNumber,isDeliveryRisk,deliveryRiskCreator,deliveryRiskCreateTime,deliveryRiskRemark"
    }
  }
}

export function buildExportBody(type) {
  let params = {
    ...defaultParams
  }

  if (type === 'total') {
    params.paramBlock.data.orderApproveTimeSection =
      getDateRange('year')

    params.paramBlock.data.purchaseOrgList = [
      '7129',
      '7107'
    ]
  } else {
    params.paramBlock.data[type] =
      getDateRange('month')

    params.paramBlock.data.purchaseOrgList = [
      '7129',
      '7107',
      '7680'
    ]
  }

  params.paramBlock.blockId = 'paramBlock'

  return {
    blocks: params
  }
}