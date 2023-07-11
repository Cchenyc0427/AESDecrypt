package main;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.security.Key;
import java.util.Base64;
import javax.crypto.Cipher;
import javax.crypto.spec.IvParameterSpec;
import javax.crypto.spec.SecretKeySpec;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AESCbcTest {
    private static final String ALGORITHM = "AES/CBC/PKCS5Padding";
    //密钥
    private static final String SECRET_KEY = "activity$Storage";
    //偏移量
    private static final String IV_PARAMETER = "activity$Storage";

    public static void main(String[] args) throws Exception {
        //需要解密的文件地址
        String filePath = "D:\\Spring_WorkSpace\\test\\src\\test\\java\\com\\example\\test\\白夜极光解密.xlsx";
        //需要解密的列数
        int columnIndex = 18; // the index of the column to decrypt (starting from 0)

        // Load the Excel file
        InputStream inputStream = new FileInputStream(filePath);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet xssfSheet =  workbook.getSheetAt(0);

        // Iterate over the rows and decrypt the specified column
        for (int k = 1; k < xssfSheet.getLastRowNum(); k++) {
            XSSFRow aRow = xssfSheet.getRow(k);
            Cell cell = aRow.getCell(columnIndex);
            if (cell != null) {
                String encryptedValue = cell.getStringCellValue();

                String decryptedValue = decrypt(encryptedValue);
                System.out.println(decryptedValue);
                cell.setCellValue(decryptedValue);


            }
        }
        try {
            //重新写入一份excel中，如果和原文件地址相同则会对原文件进行覆盖
            FileOutputStream outputStream = new FileOutputStream("D:\\Spring_WorkSpace\\test\\src\\test\\java\\com\\example\\test\\白夜极光解密.xlsx");
            workbook.write(outputStream);
            outputStream.close();
            System.out.println("数据写入成功！");
        } catch (IOException e) {
            e.printStackTrace();
        }

        workbook.close();
        inputStream.close();
    }

    private static String decrypt(String encryptedValue) throws Exception {
        byte[] secretKeyBytes = SECRET_KEY.getBytes("UTF-8");
        byte[] ivParameterBytes = IV_PARAMETER.getBytes("UTF-8");
        byte[] decodedBytes = Base64.getDecoder().decode(encryptedValue);
        Key secretKey = new SecretKeySpec(secretKeyBytes, "AES");
        IvParameterSpec ivParameterSpec = new IvParameterSpec(ivParameterBytes);

        Cipher cipher = Cipher.getInstance(ALGORITHM);
        cipher.init(Cipher.DECRYPT_MODE, secretKey, ivParameterSpec);

        byte[] decryptedValueBytes = cipher.doFinal(decodedBytes);
        return new String(decryptedValueBytes, "UTF-8");
    }

}
