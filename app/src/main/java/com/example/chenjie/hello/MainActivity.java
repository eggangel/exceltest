package com.example.chenjie.hello;

import android.Manifest;
import android.content.pm.PackageManager;
import android.os.Environment;
import android.support.v4.app.ActivityCompat;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.webkit.WebView;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;

public class MainActivity extends AppCompatActivity {
    private WebView wv_hello;
    private static String TAG = MainActivity.class.getSimpleName();
    private String captureFilePath = "/sdcard";
    private static final int REQUEST_EXTERNAL_STORAGE = 1;
    private static String[] PERMISSIONS_STORAGE = {
    Manifest.permission.READ_EXTERNAL_STORAGE,
    Manifest.permission.WRITE_EXTERNAL_STORAGE };

    public static void verifyStoragePermissions(AppCompatActivity activity) {
        int permission = ActivityCompat.checkSelfPermission(activity,Manifest.permission.WRITE_EXTERNAL_STORAGE);
        if (permission != PackageManager.PERMISSION_GRANTED) {
            ActivityCompat.requestPermissions(activity, PERMISSIONS_STORAGE,REQUEST_EXTERNAL_STORAGE);
        }
    }

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        captureFilePath = Environment.getExternalStorageDirectory().getPath();
        ArrayList list;
        TextView textView = (TextView) findViewById(R.id.tv_text);

        try {
            verifyStoragePermissions(MainActivity.this);
            list = getPointColumnData(0,0,"test.xls");
            for(int i=0;i<list.size();i++){
                textView.append(list.get(i).toString()+" ");
            }
        } catch (FileNotFoundException e) {
            Toast.makeText(MainActivity.this,"不存在该文件",Toast.LENGTH_SHORT).show();
            e.printStackTrace();
        }
    }

    public ArrayList getPointColumnData(int sheetNum, int columnIndex,
                                        String fileName) throws FileNotFoundException {
        ArrayList list = new ArrayList();
        try {
            FileInputStream inputStream = new FileInputStream(captureFilePath
                    + File.separator + fileName);
            POIFSFileSystem fSystem = new POIFSFileSystem(inputStream);
            HSSFWorkbook wb1 = new HSSFWorkbook(fSystem);
            HSSFSheet sheet = wb1.getSheetAt(sheetNum);
            int rows = sheet.getPhysicalNumberOfRows();// 获得总行数
            int columnNum = sheet.getRow(0).getPhysicalNumberOfCells();// 获得总列数
            if (rows > 0) {
                sheet.getMargin(HSSFSheet.TopMargin);
                for (int j = 0; j < rows; j++) {
                    HSSFRow row = sheet.getRow(j);
                    HSSFCell cell = row.getCell((short) columnIndex);
                    String latitude;
                    switch (cell.getCellTypeEnum()) {
                        case NUMERIC:
                            latitude = cell.getNumericCellValue() + "";
                            if (latitude == null) {

                            } else if (latitude.equals("")) {

                            } else {
                                list.add(latitude);
                            }
                            break;
                        case STRING:
                            latitude = cell.getStringCellValue();

                            Log.e(TAG, "String");
                            if (latitude == null) {

                            } else if (latitude.equals("")) {

                            } else {
                                list.add(latitude);
                            }
                            break;
                        case BOOLEAN:

                            break;

                        case FORMULA:
                            Log.e(TAG, cell.getCellFormula() + "formula");
                            break;
                        default:
                            Log.e(TAG, "unsupported cell type");
                            break;
                    }
                }
            }
            inputStream.close();
        } catch (Exception e) {

            e.printStackTrace();
        }
        return list;

    }
}
