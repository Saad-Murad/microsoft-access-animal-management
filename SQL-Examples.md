# SQL Examples

This section contains SQL examples demonstrating the database operations implemented during the development of the project.

## Filtering Records

The following query retrieves customers from a specific country:

```sql
SELECT Kundennummer, Firma, Land
FROM tblKunden
WHERE Land = "Italien";
```

## Filtering with Multiple Conditions

```sql
SELECT Kundennummer, Firma, Land
FROM tblKunden
WHERE Land = "Italien"
   OR Land = "Spanien";
```

## Pattern Matching

The `LIKE` operator can be used to find records matching a specific pattern.

```sql
SELECT Kundennummer, Firma, Land
FROM tblKunden
WHERE Firma LIKE "M*";
```

## Sorting Records

```sql
SELECT Kundennummer, Firma, Land
FROM tblKunden
ORDER BY Land, Firma;
```

## Counting Records

```sql
SELECT Count(*) AS AnzahlBestellungen
FROM tblBestellungen;
```

## Grouping Records

The following query counts the number of orders for each customer.

```sql
SELECT Kundennummer,
       Count(*) AS AnzahlvonBestellungen
FROM tblBestellungen
GROUP BY Kundennummer;
```

## Joining Related Tables

An `INNER JOIN` combines related records from the customer and order tables.

```sql
SELECT tblKunden.Kundennummer,
       tblKunden.Firma,
       Count(tblBestellungen.Kundennummer) AS AnzahlvonBestellungen
FROM tblKunden
INNER JOIN tblBestellungen
    ON tblKunden.Kundennummer = tblBestellungen.Kundennummer
GROUP BY tblKunden.Kundennummer,
         tblKunden.Firma;
```

## Appending Records

```sql
INSERT INTO tblKunden
    (Firma, Ansprechpartner, Position, Straße, PLZ, Ort, Land, Telefon, Telefax)
SELECT Firma, Ansprechpartner, Position, Straße, PLZ, Ort, Land, Telefon, Telefax
FROM tblKunden_aus_Bayern;
```

## Deleting Records

```sql
DELETE FROM tblKunden
WHERE Firma = "Bei Edmund"
   OR Firma = "Zum Bullen";
```

## SQL Concepts Used

The database work includes practical use of:

- `SELECT`
- `FROM`
- `WHERE`
- `AND` / `OR`
- `LIKE`
- `ORDER BY`
- `GROUP BY`
- `COUNT`
- `INNER JOIN`
- `INSERT INTO`
- `DELETE`
