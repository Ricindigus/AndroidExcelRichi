package com.example.user.androidexcel;

import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import android.app.Activity;
import android.content.Context;
import android.os.Bundle;
import android.os.Environment;
import android.support.v7.widget.LinearLayoutManager;
import android.support.v7.widget.RecyclerView;
import android.util.Log;
import android.view.View;
import android.view.View.OnClickListener;
import android.widget.Button;
import android.widget.Toast;

import static android.os.Environment.getExternalStorageDirectory;

public class MainActivity extends AppCompatActivity implements OnClickListener{

    RecyclerView recyclerView;
    LinearLayoutManager linearLayoutManager;
    MarcoAdapter marcoAdapter;
    Button writeExcelButton,readExcelButton;
    ArrayList<ItemMarco> itemMarcos;
    static String TAG = "ExelLog";

    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        itemMarcos = new ArrayList<ItemMarco>();
        recyclerView = (RecyclerView) findViewById(R.id.recycler);
        writeExcelButton = (Button) findViewById(R.id.writeExcel);
        writeExcelButton.setOnClickListener(this);
        readExcelButton = (Button) findViewById(R.id.readExcel);
        readExcelButton.setOnClickListener(this);
        linearLayoutManager = new LinearLayoutManager(getApplicationContext());
        linearLayoutManager.setOrientation(LinearLayoutManager.VERTICAL);
        recyclerView.setLayoutManager(linearLayoutManager);
        itemMarcos.add(new ItemMarco("1","2","3"));
        itemMarcos.add(new ItemMarco("1","2","3"));
        itemMarcos.add(new ItemMarco("1","2","3"));
        itemMarcos.add(new ItemMarco("1","2","3"));
        itemMarcos.add(new ItemMarco("1","2","3"));
        itemMarcos.add(new ItemMarco("1","2","3"));
        itemMarcos.add(new ItemMarco("1","2","3"));
        itemMarcos.add(new ItemMarco("1","2","3"));
        marcoAdapter = new MarcoAdapter(itemMarcos);
        recyclerView.setAdapter(marcoAdapter);
    }

    public void onClick(View v)
    {
        switch (v.getId())
        {
            case R.id.writeExcel:
//                File nuevaCarpeta = new File(getExternalStorageDirectory(), "RICARDODIR");
//                nuevaCarpeta.mkdirs();
                File wallpaperDirectory = new File("/sdcard/Wallpaper/");
// have the object build the directory structure, if needed.
                wallpaperDirectory.mkdirs();
//// create a File object for the output file
//                File outputFile = new File(wallpaperDirectory, filename);
//// now attach the OutputStream to the file object, instead of a String representation
//                FileOutputStream fos = new FileOutputStream(outputFile);
                saveExcelFile(this,"/myExcel.xls");
                break;
            case R.id.readExcel:
                itemMarcos = readExcelFile(this,"/myExcel.xls");
                marcoAdapter = new MarcoAdapter(itemMarcos);
                recyclerView.setAdapter(marcoAdapter);
                Toast.makeText(this, "llego aqui", Toast.LENGTH_SHORT).show();
                break;
        }
    }

    private boolean saveExcelFile(Context context, String fileName) {

        // check if available and not read only
        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) {
            Log.e(TAG, "Storage not available or read only");
            return false;
        }

        boolean success = false;

        //New Workbook
        Workbook wb = new HSSFWorkbook();

        //New Sheet
        Sheet sheet1 = null;
        sheet1 = wb.createSheet("myOrder");


        //Cell style for header row
        CellStyle cs = wb.createCellStyle();
        cs.setFillForegroundColor(HSSFColor.LIME.index);
        cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        //cabecera excel
        Row row0 = sheet1.createRow(0);
        Cell c1 = null;
        c1 = row0.createCell(0);
        c1.setCellValue("Item Number");
        c1.setCellStyle(cs);
        c1 = row0.createCell(1);
        c1.setCellValue("Quantity");
        c1.setCellStyle(cs);
        c1 = row0.createCell(2);
        c1.setCellValue("Price");
        c1.setCellStyle(cs);

        CellStyle cs1 = wb.createCellStyle();
        cs1.setFillForegroundColor(HSSFColor.WHITE.index);
        //Celdas Excel
        for (int i = 0; i <3 ; i++) {
            Row row = sheet1.createRow(i + 1);
            Cell c = null;

            c = row.createCell(0);
            c.setCellValue((i + 1)+"");
            c.setCellStyle(cs1);

            c = row.createCell(1);
            c.setCellValue("Cantidad" + (i+1));
            c.setCellStyle(cs1);

            c= row.createCell(2);
            c.setCellValue("Precio" + (i+1));
            c.setCellStyle(cs1);
        }


        sheet1.setColumnWidth(0, (15 * 500));
        sheet1.setColumnWidth(1, (15 * 500));
        sheet1.setColumnWidth(2, (15 * 500));

        // Create a path where we will place our List of objects on external storage
        File file = new File(context.getExternalFilesDir(null), fileName);
        FileOutputStream os = null;

        try {
            os = new FileOutputStream(file);
            wb.write(os);
            Log.w("FileUtils", "Writing file" + file);
            success = true;
        } catch (IOException e) {
            Log.w("FileUtils", "Error writing " + file, e);
        } catch (Exception e) {
            Log.w("FileUtils", "Failed to save file", e);
        } finally {
            try {
                if (null != os)
                    os.close();
            } catch (Exception ex) {
            }
        }
        return success;
    }

    private ArrayList<ItemMarco> readExcelFile(Context context, String filename) {

        ArrayList<ItemMarco> itemMarcos1 = new ArrayList<ItemMarco>();
        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) Log.e(TAG, "Storage not available or read only");
        else{
            try{
                // Creating Input Stream
                File file = new File(context.getExternalFilesDir(null), filename);
                FileInputStream myInput = new FileInputStream(file);

                InputStream stream = context.getAssets().open("myExcel.xls");
                // Create a POIFSFileSystem object
                POIFSFileSystem myFileSystem = new POIFSFileSystem(stream);

                // Create a workbook using the File System
                HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);

                // Get the first sheet from workbook
                HSSFSheet mySheet = myWorkBook.getSheetAt(0);

                /** We now need something to iterate through the cells.**/
                Iterator rowIter = mySheet.rowIterator();

                while(rowIter.hasNext()){
                    HSSFRow myRow = (HSSFRow) rowIter.next();
                    Iterator cellIter = myRow.cellIterator();
                    String m = "";
                    ItemMarco itemMarco = new ItemMarco();
                    int contador = 1;
                    while(cellIter.hasNext()){
                        HSSFCell myCell = (HSSFCell) cellIter.next();
                        Log.d(TAG, "Cell Value: " +  myCell.toString());
//                        m = m + myCell.toString();
                        if(contador == 1) itemMarco.setNumero(""+(int)myCell.getNumericCellValue());
                        if(contador == 2) itemMarco.setRuc(myCell.toString());
                        if(contador == 3) itemMarco.setRazonSocial(""+(int)myCell.getNumericCellValue());
                        contador++;
//                    Toast.makeText(context, "cell Value: " + myCell.toString(), Toast.LENGTH_SHORT).show();
                    }
                    Toast.makeText(context, "m: " + m, Toast.LENGTH_SHORT).show();
                    itemMarcos1.add(itemMarco);
                }
            }catch (Exception e){e.printStackTrace(); }

        }
        return itemMarcos1;
    }

    public static boolean isExternalStorageReadOnly() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED_READ_ONLY.equals(extStorageState)) {
            return true;
        }
        return false;
    }

    public static boolean isExternalStorageAvailable() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED.equals(extStorageState)) {
            return true;
        }
        return false;
    }
}
