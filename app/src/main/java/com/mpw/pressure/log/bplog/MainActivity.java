package com.mpw.pressure.log.bplog;

import android.os.Bundle;
import android.support.v7.app.AppCompatActivity;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.view.View.OnClickListener;
import android.widget.Toast;

import java.io.IOException;

import jxl.write.WriteException;

public class MainActivity extends AppCompatActivity implements OnClickListener {
    /**
     * Called when the activity is first created.
     */
    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        View writeExcelButton = findViewById(R.id.writeExcel);
        writeExcelButton.setOnClickListener(this);


    }

    private void writeTOXsl() {
        WriteExcel writeExcel = new WriteExcel();
        try {
            writeExcel.write("_Malith_BP_LOG.xls", MainActivity.this);
            Toast.makeText(this, "done", Toast.LENGTH_SHORT).show();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

    public void onClick(View v) {
        switch (v.getId()) {
            case R.id.writeExcel:
                writeTOXsl();
                break;
        }
    }


    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        // Inflate the menu; this adds items to the action bar if it is present.
        getMenuInflater().inflate(R.menu.menu_main, menu);
        MenuItem searchItem = menu.findItem(R.id.menu_search);

        return true;
    }

    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        // Handle action bar item clicks here. The action bar will
        // automatically handle clicks on the Home/Up button, so long
        // as you specify a parent activity in AndroidManifest.xml.
        switch (item.getItemId()) {
            case R.id.menu_export:

                return true;

            default:
                return super.onOptionsItemSelected(item);
        }
    }
}
