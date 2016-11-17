# salesforce-postcodes-pl-converter-vba-msaccess

The VBA code working only for NNNNN postcodes, where N is digit between 0 and 9

1. create ms access database
2. create first table:
- name: DataLoader
- field name: MASTERLABEL ; field type: Text ; field size: 255
- field name: OPERATION ; field type: Text ; field size: 255
- field name: VALUE ; field type: Memo
3. create second table:
- name: Structure
- field name: Territory ; field type: Text ; field size: 255
- field name: PoBox ; field type: Text ; field size: 255
4. create module and paste VBA code
5. upload data to DataLoader table (source details in file: SalesForceDataLoader.png) and run VBA code
