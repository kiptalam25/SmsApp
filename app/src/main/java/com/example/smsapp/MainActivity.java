package com.example.smsapp;

import androidx.annotation.NonNull;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;

import android.Manifest;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.telephony.SmsManager;
import android.util.Log;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class MainActivity extends AppCompatActivity {
    // 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      token
    private static final int PERMISSION_SEND_SMS = 123;
    private static final String TAG = "Tag";
    //    private static CharArrayWriter workbook;
    Button btn_send,btn_contacts;
    EditText txt_message,txt_phone;
    List<Contact> contacts =new ArrayList<>();
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        txt_message=findViewById(R.id.txt_message);
        btn_send=findViewById(R.id.btn_send);
        txt_phone=findViewById(R.id.txt_phone);
        btn_contacts=findViewById(R.id.btn_contacts);
        btn_contacts.setOnClickListener(v->{
           startActivity(new Intent(MainActivity.this,Contacts.class));
        });
        newWork();
        btn_send.setOnClickListener(v-> {
int x=0;
            for (int i = 0; i < 100; i++){

            if(i==x) {
                x+=6;
                try {
                    Thread.sleep(10000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            }
                if (ContextCompat.checkSelfPermission(
                        MainActivity.this,
                        Manifest.permission.SEND_SMS) == PackageManager.PERMISSION_GRANTED) {
                    String message = txt_message.getText().toString().trim();
                    String modified = message;

                sendMessage(modified);
//                    readExcelFromStorage(MainActivity.this, "workbook.xls");
                } else {
                    ActivityCompat.requestPermissions(MainActivity.this,
                            new String[]{Manifest.permission.SEND_SMS},
                            PERMISSION_SEND_SMS);
                }

        }
        });


    }

    private void sendMessage(String message) {
        String phone=txt_phone.getText().toString().trim();
        SmsManager smsManager=SmsManager.getDefault();
        smsManager.sendTextMessage(phone,null,message,null,null);
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);

        if(requestCode==100 && grantResults.length> 0 && grantResults[0]== PackageManager.PERMISSION_GRANTED){
//            sendMessage();
        }else {
            Toast.makeText(this, "Permission Denied", Toast.LENGTH_SHORT).show();
        }
    }


    private Sheet sheet = null;

    private Cell cell = null;
    private static String EXCEL_SHEET_NAME = "Sheet1";
    Workbook workbook = new HSSFWorkbook();


    public void newWork(){


// Create a new sheet in a Workbook and assign a name to it
        sheet = workbook.createSheet(EXCEL_SHEET_NAME);
        Row row = sheet.createRow(0);
        Row row1 = sheet.createRow(1);

// Cell style for a cell
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillForegroundColor(HSSFColor.AQUA.index);
        cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);

// Creating a cell and assigning it to a row
        cell = row.createCell(0);
        Cell cell1=row1.createCell(0);
        cell1.setCellValue("joseph");
        cell1.setCellStyle(cellStyle);
         cell1=row1.createCell(1);
        cell1.setCellValue(4567894);
        cell1.setCellStyle(cellStyle);
// Setting Value and Style to the cell
        cell.setCellValue("Kiptalam");
        cell.setCellStyle(cellStyle);
        cell = row.createCell(1);
        cell.setCellValue(4567890);
        cell.setCellStyle(cellStyle);

        storeExcelInStorage(MainActivity.this,"workbook.xls");

    }

    public static void readExcelFromStorage(Context context, String fileName) {

        List<Contact> contacts =new ArrayList<>();

        File file = new File(context.getExternalFilesDir(null), fileName);
        FileInputStream fileInputStream = null;

        try {
            fileInputStream = new FileInputStream(file);
            System.out.println("--------------------------------------Reading from Excel" + file);

            // Create instance having reference to .xls file
            Workbook workbook = new HSSFWorkbook(fileInputStream);

            // Fetch sheet at position 'i' from the workbook
            System.out.println("------------------------------"+workbook.getSheetName(0));
            Sheet sheet = workbook.getSheetAt(0);
            System.out.println("------------------------------"+sheet.getPhysicalNumberOfRows());
            // Iterate through each row
            for (Row row : sheet) {
                Contact contact=new Contact();

                System.out.println("------------------------------"+row.getRowNum());

//                if (row.getRowNum() >= 0) {

                    // Iterate through all the cells in a row (Excluding header row)
                    Iterator<Cell> cellIterator = row.cellIterator();

                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        // Check cell type and format accordingly

                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_NUMERIC:
                                // Print cell value
                                contact.setPhoneNo(String.valueOf(cell.getNumericCellValue()));
                                System.out.println("-----------------------"+cell.getNumericCellValue());
                                break;

                            case Cell.CELL_TYPE_STRING:
                                contact.setName(cell.getStringCellValue());
                                System.out.println("--------------------------"+cell.getStringCellValue());
                                break;

                        }

                    }

//                }
                System.out.println("---------- Adding contact");
                contacts.add(contact);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            try {
                if (null != fileInputStream) {
                    fileInputStream.close();
                }
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
    }

    private boolean  storeExcelInStorage(Context context, String fileName) {

        System.out.println("-----------------------------------"+ "saved");
        boolean isSuccess;


        File file = new File(context.getExternalFilesDir(null), fileName);
        FileOutputStream fileOutputStream = null;

        try {
            fileOutputStream = new FileOutputStream(file);
//            Workbook wb=new HSSFWorkbook(fileOutputStream);
            workbook.write(fileOutputStream);
            Log.e(TAG, "Writing file" + file);
            isSuccess = true;
        } catch (IOException e) {
            Log.e(TAG, "Error writing Exception: ", e);
            isSuccess = false;
        } catch (Exception e) {
            Log.e(TAG, "Failed to save file due to Exception: ", e);
            isSuccess = false;
        } finally {
            try {
                if (null != fileOutputStream) {
                    fileOutputStream.close();
                }
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
        return isSuccess;
    }
}