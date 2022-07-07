package com.example.smsapp;

import androidx.annotation.NonNull;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import android.annotation.SuppressLint;
import android.content.ContentResolver;
import android.content.pm.PackageManager;
import android.database.Cursor;
import android.os.Bundle;
import android.provider.ContactsContract;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.ListView;

import org.json.JSONObject;

import java.util.ArrayList;
import java.util.List;

public class Contacts extends AppCompatActivity {
    public static final int REQUEST_READ_CONTACTS = 79;
    ListView list;
    ArrayList mobileArray=new ArrayList();
    Button btn_request;
    List<Contact> contacts=new ArrayList<>();
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_contacts);
        list = findViewById(R.id.list);
        btn_request = findViewById(R.id.btn_request);

        btn_request.setOnClickListener(v -> {

        });
        if (ActivityCompat.checkSelfPermission(this, android.Manifest.permission.READ_CONTACTS)
                == PackageManager.PERMISSION_GRANTED) {
            contacts = getAllContacts();
            System.out.println( contacts.toArray());
            for(Contact c:contacts){
                mobileArray.add(c.name+"\n"+c.phoneNo);
            }


            ArrayAdapter adapter = new ArrayAdapter<String>(this,
                    android.R.layout.simple_list_item_1, android.R.id.text1, mobileArray);
            list.setAdapter(adapter);

        } else {
            requestPermission();
        }


    }

    private void requestPermission() {
        if (ActivityCompat.shouldShowRequestPermissionRationale(this, android.Manifest.permission.READ_CONTACTS)) {
            // show UI part if you want here to show some rationale !!!
        } else {
            ActivityCompat.requestPermissions(this, new String[]{android.Manifest.permission.READ_CONTACTS},
                    REQUEST_READ_CONTACTS);
        }
        if (ActivityCompat.shouldShowRequestPermissionRationale(this, android.Manifest.permission.READ_CONTACTS)) {
        } else {
            ActivityCompat.requestPermissions(this, new String[]{android.Manifest.permission.READ_CONTACTS},
                    REQUEST_READ_CONTACTS);
        }
    }

    @SuppressLint("Range")
    private List<Contact> getAllContacts() {
        List<Contact> contacts1 = new ArrayList<>();
        ContentResolver cr = getContentResolver();
        Cursor cur = cr.query(ContactsContract.Contacts.CONTENT_URI,
                null, null, null, null);


        if ((cur != null ? cur.getCount() : 0) > 0)
            while (cur != null && cur.moveToNext()) {
                String id = cur.getString(
                        cur.getColumnIndex(ContactsContract.Contacts._ID));
                String name = cur.getString(cur.getColumnIndex(
                        ContactsContract.Contacts.DISPLAY_NAME));
                JSONObject obj = new JSONObject();
//                nameList.add(name);
                if (cur.getInt(cur.getColumnIndex(ContactsContract.Contacts.HAS_PHONE_NUMBER)) > 0) {
                    Cursor pCur = cr.query(
                            ContactsContract.CommonDataKinds.Phone.CONTENT_URI,
                            null,
                            ContactsContract.CommonDataKinds.Phone.CONTACT_ID + " = ?",
                            new String[]{id}, null);
                    while (pCur.moveToNext()) {
                        String phoneNo = pCur.getString(pCur.getColumnIndex(
                                ContactsContract.CommonDataKinds.Phone.NUMBER));
                        String nam = pCur.getString(pCur.getColumnIndex(ContactsContract.CommonDataKinds.Phone.DISPLAY_NAME));
                        try {
                            Contact contact=new Contact(nam,phoneNo);
                            contacts1.add(contact);

                        } catch (Exception e) {
                            System.out.println(e.getMessage());
                        }
//                        nameList.add(contacts);
                    }
                    pCur.close();
                }
            }

        if (cur != null) {
            cur.close();
        }
        return contacts1;
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);
        switch (requestCode) {
            case REQUEST_READ_CONTACTS: {
                if (grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                    contacts = getAllContacts();
                } else {
                    // permission denied,Disable the
                    // functionality that depends on this permission.
                }
                return;
            }
        }

    }
}