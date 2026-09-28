# VBA Examples

This section contains selected VBA implementations used in the Microsoft Access Animal Management System.

## Maximizing a Form

The main animal inventory form is automatically maximized when it opens.

```vba
Private Sub Form_Open(Cancel As Integer)
    DoCmd.Maximize
End Sub
```

## Closing a Form

A command button closes the current form.

```vba
Private Sub btnSchließen_Click()
    DoCmd.Close acForm, Me.Name
End Sub
```

## Filtering Animals by Location

The selected location is used to filter the animal records displayed in the form.

```vba
Private Sub cmbStandort_AfterUpdate()

    Me.Filter = "Standort_ID = " & Me.cmbStandort
    Me.FilterOn = True

End Sub
```

## Removing a Filter

All animal records can be displayed again by disabling the active filter.

```vba
Private Sub optAlleTiere_GotFocus()
    Me.FilterOn = False
End Sub
```

## Opening Details for the Selected Animal

The details form opens only the animal currently selected by the user.

```vba
Private Sub btnTierdetails_Click()

    DoCmd.OpenForm "frmTierdetails", , , _
        "Tier_ID = " & Me.Tier_ID

End Sub
```

## Filtering by Food Requirement

Users can specify the maximum amount of food an animal should consume per day.

```vba
Private Sub cmdAnzeigen_Click()

    Me.Filter = "Futtermenge <= " & Me.txtFuttermenge
    Me.FilterOn = True

End Sub
```

## Counting Animals

The number of animals belonging to the selected location is calculated and displayed in a text field.

```vba
Private Sub btnZählen_Click()

    Me.txtAnzahl = DCount("*", "tblTier", _
        "Standort_ID = " & Me.cmbTierheim)

End Sub
```

## Opening Animals from the Same Subcategory

The application can open another form and display animals belonging to the same subcategory.

```vba
Private Sub btnAndereTiere_Click()

    DoCmd.OpenForm "frmAndere Tiere", , , _
        "Unterkategorie_ID = " & Me.txtUnterkategorieID

End Sub
```
