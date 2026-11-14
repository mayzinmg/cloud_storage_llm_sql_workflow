/* B2 TAKASHIMAYA FOOD HALL, CENTRAL BAZAAR
   Gourmet & Groceries
   1/11/26–30/11/26
   Setup: 31/10/26
   Class: 581925
*/
DECLARE @CompanyId int = 64;
DECLARE @Ref3 varchar(20) = '581925';
DECLARE @Ref4 varchar(20) = '58';

DECLARE @Map TABLE
(
    TerminalNo varchar(50)   NOT NULL PRIMARY KEY,
    Brand      nvarchar(200) NOT NULL,
    ACNo       varchar(50)   NOT NULL,
    DKU        varchar(50)   NOT NULL
);

INSERT INTO @Map (TerminalNo, Brand, ACNo, DKU)
VALUES
    ('79313', N'Purest Co',     '505184', '10460137'),
    ('79323', N'KELE',          '504183', '10375905'),
    ('79333', N'Qi Cha',        '505788', '10507960'),
    ('79343', N'Cheerful Goat', '505620', '10494904');

/* 1) Location */
UPDATE l
SET
    l.LocationDesc = m.Brand,
    l.ref7         = m.ACNo
FROM location l
INNER JOIN posmachines p ON p.locationid = l.locationid
INNER JOIN @Map m        ON m.TerminalNo = p.posid
WHERE l.companyid = @CompanyId;

/* 2) Category */
UPDATE c
SET
    c.categorydesc = m.DKU,
    c.no           = m.DKU,
    c.ref3         = @Ref3,
    c.ref4         = @Ref4,
    c.ref5         = m.ACNo
FROM category c
INNER JOIN posmachines p ON p.locationid = c.locationid
INNER JOIN location l    ON l.locationid = p.locationid
INNER JOIN @Map m        ON m.TerminalNo = p.posid
WHERE l.companyid = @CompanyId;

/* 3) POSMachines */
UPDATE p
SET
    p.tenantno = m.ACNo
FROM posmachines p
INNER JOIN location l ON l.locationid = p.locationid
INNER JOIN @Map m     ON m.TerminalNo = p.posid
WHERE l.companyid = @CompanyId;

/* 4) POSGroup */
UPDATE po
SET
    po.POSGroupName = m.Brand
FROM POSGroup po
INNER JOIN posmachines p ON p.POSGroupId = po.POSGroupId
INNER JOIN location l    ON l.locationid = p.LocationId
INNER JOIN @Map m        ON m.TerminalNo = p.posid
WHERE l.companyid = @CompanyId;