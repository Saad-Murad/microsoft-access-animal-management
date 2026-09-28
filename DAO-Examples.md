# DAO and Recordset Examples

This section contains examples of working with Microsoft Access databases programmatically using DAO (Data Access Objects).

## Working with a Database and Recordset

A DAO `Database` object provides access to the current Access database, while a `Recordset` represents a set of records that can be processed with VBA.

```vba
Dim db As DAO.Database
Dim rs As DAO.Recordset

Set db = CurrentDb
Set rs = db.OpenRecordset("tblKunden")
```

## Counting Records

The following procedure opens a table and determines the number of records.

```vba
Sub AnzahlKunden()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblKunden")

    rs.MoveLast

    MsgBox rs.RecordCount

End Sub
```

## Navigating Through Records

DAO provides methods such as `MoveFirst`, `MoveNext`, `MovePrevious`, and `MoveLast` for navigating through a recordset.

The following example processes records until the end of the recordset is reached:

```vba
Sub KundenAusDeutschland()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("qry Kunden aus Deutschland")

    rs.MoveFirst

    Do Until rs.EOF
        MsgBox rs!Firma
        rs.MoveNext
    Loop

End Sub
```

## Editing Records

Existing records can be modified using `Edit` and `Update`.

```vba
Sub UpdateCountry()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblKunden")

    rs.MoveFirst

    Do Until rs.EOF

        If rs!Land = "Großbritannien" Then

            rs.Edit
            rs!Land = "UK"
            rs.Update

        End If

        rs.MoveNext

    Loop

End Sub
```

## DAO Concepts Used

The database exercises include practical use of:

- `DAO.Database`
- `DAO.Recordset`
- `CurrentDb`
- `OpenRecordset`
- `RecordCount`
- `MoveFirst`
- `MoveNext`
- `MoveLast`
- `EOF`
- `Edit`
- `Update`
- Recordset field access

## Summary

I implemented DAO-based procedures to access, navigate, count, read, and modify records directly through VBA.

These examples demonstrate programmatic database operations beyond the standard Microsoft Access form interface.
