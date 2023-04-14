# Search like "Google" in Microsoft Access form

By using query, form and sub-form you can get a "Google" like search.
Where the form update data for each letter you enter.

Lets say you have a product list you want to search in.
Itemno, EAN code, description, price.

Create a query with all the fields you want to show in the list/form.
In this example we want to show all fields in the list and we add a "search" column at the end.
We will use this "search" column to filter the data that will show at the sub-form.
See SQL sample below

```sql
SELECT ProductTable.*, [itemno] & [ean] & [description1] & [unitprice] & [costprice] AS Search
FROM ProductTable
WHERE ((([itemno] & [ean] & [description1] & [unitprice] & [costprice]) Like "*" & [Froms]![Main]![txtDummy] & "*"));
```

You can also use "Create query" in Microsoft Access if you do not want to write the SQL statment.

As you see in the SQL statment we select all field from a table. And then we create a "search" column with all the columns we want the search to contain.
In this sample we can search on columns: Itemno, ean, Description1, Unitprice and costprice.
SO for each letter you enter in the textbox for search will then check these columns if the search match the criteria and show the result.
Hide txtDummy by make the background and text color the same as background color of the form.

* Create a form with 2 textboxes named txtSearch and txtDummy
* Create a subform and link the source to the form from the SQL statment.

Add VBA code to txtSeach textbox change event.

```vba
Private Sub txtSearch_Change()
Me.txtDummy = Me.txtSearch.Text
Me.Q_ProductSearch_subform.Requery
End Sub
```

!([Textbox txtDummy](https://github.com/Idemar/Programming/blob/master/Microsoft%20Access/Search%20like%20Google%20in%20form/images/Skjema_hoved.png))

!([Before filter in the search box](https://github.com/Idemar/Programming/blob/master/Microsoft%20Access/Search%20like%20Google%20in%20form/images/Hovedskjema.png))

!([After filter in the search box](https://github.com/Idemar/Programming/blob/master/Microsoft%20Access/Search%20like%20Google%20in%20form/images/Hovedskjema2.png))
