package com.swinkpay.reportautomation.SwinkpayReportsAutomation;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class Automate {
        static final int merchantColumnNumber = 2;
        static final int responseCodeColumnNumber = 19;
        static final int transactionStatusColumnNumber = 16;

        static final int lastSuccessfulTxnTime = 17;

        static  final int skipInitialNumOfRows = 3;

        static final int paymentGatewayColumnNumber = 21;

        static final int qrCodeTypeColumnNumber = 25;

        public static String generateReport() throws IOException {

            //obtaining i/p bytes from file
            FileInputStream fis = new FileInputStream(new File("C:\\Users\\purnima\\Downloads\\Automationn.xls"));
            HSSFWorkbook wb = new HSSFWorkbook(fis);
            //creating a Sheet object to retrieve the object
            HSSFSheet sheet = wb.getSheetAt(0);
            FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
            int rowNumber = 0;
            Map<String, Merchant> resultMap = new HashMap<>();
            for (Row row : sheet) //iteration of all rows
            {
                if (rowNumber > skipInitialNumOfRows && rowNumber!= sheet.getLastRowNum()) {


                    Merchant merchantObj = null;
                    String merchantName = row.getCell(merchantColumnNumber).getStringCellValue();
                    String txnStatus = row.getCell(transactionStatusColumnNumber).getStringCellValue();
                    String responseCode = row.getCell(responseCodeColumnNumber).getStringCellValue();
                    String paymentGateway = row.getCell(paymentGatewayColumnNumber).getStringCellValue();
                    String qrCodeType = row.getCell(qrCodeTypeColumnNumber).getStringCellValue();

                    if (resultMap.containsKey(merchantName)) {
                        merchantObj = resultMap.get(merchantName);
                    } else {
                        merchantObj = new Merchant();
                        merchantObj.setMerchantName(merchantName);
                    }
                    merchantObj.getResponseCodeMap().put(responseCode, merchantObj.getResponseCodeMap().getOrDefault(responseCode, 0l) + 1);
                    merchantObj.getTxnStatusMap().put(txnStatus, merchantObj.getTxnStatusMap().getOrDefault(txnStatus, 0l) + 1);
                    merchantObj.getPaymentTypeMap().put(paymentGateway + "-" + qrCodeType, merchantObj.getPaymentTypeMap().getOrDefault(paymentGateway + "-" + qrCodeType, 0L) + 1);
                    merchantObj.setOverallCount(merchantObj.getOverallCount() + 1);

                    resultMap.put(merchantName, merchantObj);
                }
                rowNumber = rowNumber + 1;
            } //END

            System.out.println(" Last Successful Txn Time Happened At "+sheet.getRow(sheet.getLastRowNum()-1).getCell(lastSuccessfulTxnTime).getStringCellValue()+"\n");

            System.out.println("<------------------------------------------------------REPORT" +
                    "------------------------------------------------------------------------------------->");

            for(Map.Entry<String, Merchant> entry:resultMap.entrySet()){
                System.out.println(entry+"\n");
                System.out.println("----------------------------------------------------------");
            }

            return resultMap.toString();
        }
    }

    class  Merchant {
        String merchantName;
        double overallCount;
        Map<String, Long> responseCodeMap;
        Map<String, Long> txnStatusMap;
        Map<String, Long> paymentTypeMap;

        public Merchant() {
            overallCount = 0;
            responseCodeMap = new HashMap<>();
            txnStatusMap = new HashMap<>();
            paymentTypeMap = new HashMap<>();
        }

        //Getters and setters

/*        public String getMerchantName() {
        return merchantName;
        }*/

        public void setMerchantName(String merchantName) {
            this.merchantName = merchantName;
        }

        public double getOverallCount() {
            return overallCount;
        }

        public void setOverallCount(double overallCount) {
            this.overallCount = overallCount;
        }

        public Map<String, Long> getResponseCodeMap() {
            return responseCodeMap;
        }

        public void setResponseCodeMap(Map<String, Long> responseCodeMap) {
            this.responseCodeMap = responseCodeMap;
        }

        public Map<String, Long> getTxnStatusMap() {
            return txnStatusMap;
        }

        public void setTxnStatusMap(Map<String, Long> txnStatusMap) {
            this.txnStatusMap = txnStatusMap;
        }

        public Map<String, Long> getPaymentTypeMap() {
            return paymentTypeMap;
        }

        public void setPaymentTypeMap(Map<String, Long> paymentTypeMap) {
            this.paymentTypeMap = paymentTypeMap;
        }

        @Override

        public String toString() {
            return "Merchant{" +
                    "merchantName=' " + merchantName + '\'' +
                    ",overallCount=" + overallCount +
                    ",responseCodeMap=" + responseCodeMap +
                    ", txnStatusMap=" + txnStatusMap +
                    ",paymentTypeMap=" + paymentTypeMap +
                    '}';
        }
    }
