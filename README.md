# Summary
 XML Generator to assist in the creation and modification of tags for Ovarro's TBOX RTU

## Requires
TwinSoft [RTU](https://www.servelectechnologies.com/servelec-technologies/products-services/remote-telemetry-units/tbox/tbox-lt2/) software from Ovarro 
TwinSoftImEx Tool from Ovarra

## Why
For larger projects, the creation of tags is a tedious and error prone process. Fortunetly, Ovarra' TwinSoftImEx exports tags in XML format that can be modified in Excel and then imported back. Although faster than using the Twinsoft UI to individually create tags, editing XML is still error prone.

tag_importer used templates in created in Excel from which to seed the generation of a tags XML file compatable with TwinSoftImEx.

## Use Cases
#1 An engineer must create time control tags for serveral root tags. Example 3 Tags shown below.

|ROOT TAG|	DESCRIPTION	| 
|--------|----------|
|KM_103 |CHAMBER 1 MOTHER FEED SCHEDULE  
|KV_103 |CHAMBER 1 VEG FEED SCHEDULE  
|KF_103 |CHAMBER 1 CLONE FEED SCHEDULE  

#2 An engineer need to clone a folder to another and resulting an new tags and descriptions and modbus addresses that mirror those in the folder AND having a unique tagname.

|TAG|	DESCRIPTION	| MODBUS ADDRESS|FOLDER|
|---|---|---|---|
|LI_101| CHAMBER 1 H2O INTAKE TANK LEVEL | 1000|CHAMBER 1
|TI_101| CHAMBER 1 INTAKE TANK H2O TEMPERATURE | 1002|CHAMBER 1
|TIC_101_SP|CHAMBER 1TANK H2O TEMPERATURE CONTROLLER SP| 1100| CHAMBER 1\SOFTS

clone folder ends up with the following newtags
|TAG|	DESCRIPTION	| MODBUS ADDRESS|FOLDER|
|---|---|---|---|
|LI_201| CHABMER 2 H2O INTAKE TANK LEVEL | 2000|CHAMBER 2
|TI_201| CHABMER 2INTAKE TANK H2O TEMPERATURE | 2002|CHAMBER 1
|TIC_201_SP|CHABMER 2TANK H2O TEMPERATURE CONTROLLER SP| 2100| CHAMBER 2\SOFTS


## How

Each root tag ends up with a SUFFIX appended to it as well as the DESCRIPTION expanded to include the SUFFIX description. It does not have to be suffix and can be a prefix or something in the middle.

An Excel sheet called TEMPLATE contains the rules used to seed the tag generation process. A sample for a TIME control template is shown below.

|TEMPLATE|	SUFFIX	|DESCRIPTION|	TYPE|	INITIAL_VALUE|	SCRIPT_VALUE|
|--------|----------|-----------|-------|----------------|--------------|
|TIME|	_OFAP|	OFF AM/PM (1=PM)|	UINT16|	-9999|	1
|TIME|	_OFH|	OFF HR	|UINT16|	-9999|	16
|TIME|	_OFM|	OFF MIN	|UINT16|	-9999|	16
|TIME|	_ONAP|	ON AM/PM (1=PM)|	UINT16|	-9999|	0
|TIME|	_ONH|	ON HR	|UINT16|	-9999|	7
|TIME|	_ONM|	ON MIN	|UINT16|	-9999|	0


_Constraint: all columns must contain values. A check is made at runtime to ensure this contraint is met.__
Column descriptions:

|COLUMN|DESCRIPTIION              |
|------|--------------------------|
| TEMPLATE| The name of the template
| SUFFIX| what to insert/append to the root tag
| DESCRIPTION| what to insert/append to the root tag description
| TYPE | tag data type BOOL, UINT8, UINT16,INT17, UINT32, INT32, FLOAT
| INITIAL_VALUE | tags initial value, if none enter -9999
| SCRIPT_VALUE | tag value in auto script generation if none enter -9999

In the same Excel file, a TAGS sheet is created where each row defines a template instance

|CLASS|TAG_NAME|TAG_PATTERN|DESCRIPTION|TEMPLATE|FOLDER|GROUP|
|-----|--------|-----------|-----------|--------|------|-----|
|GENERATE||KM_103*|CHAMBER 1 MOTHER FEED SCHEDULE *|TIME|CHAMBER 1 SOFTS|CHAMBER 1
|GENERATE||KV_103*|CHAMBER 1 VEG FEED SCHEDULE *|TIME|CHAMBER 1 SOFTS|CHAMBER 1
|GENERATE||KC_103*|CHAMBER 1 CLONE FEED SCHEDULE *|TIME|CHAMBER 1 SOFTS|CHAMBER 1


Column descriptions:
|COLUMN|DESCRIPTIION              |
|------|--------------------------|
| GENERATE| == GENERATE for template instances
| TAG_NAME | empty for template instances
| TAG_PATTERN | Root tag name and an * placed where the template suffix need is to be inserted
| DESCRIPTION | Root tag description and an * placed where the template description is to be inserted
| TEMPLATE | template name in the TEMPLATE sheet
| FOLDER | folder in twinsoft where the generated tags will be stored under
| GROUP | Memory map entry in MEMORY_MAP sheet 

Lastly, the same EXCEL file must contain a MEMORY_MAP sheet specifying the modbus addresses for the various data types and groups


|GROUP|	FORMAT|	START_ADDRESS|	LENGTH|	TS_FORMAT|	TS_FORMAT|	TS_SIGNED|
|-------|-------|---|---|---|---|---|
|CHAMBER 1|	FLOAT|	1000|	100|	1|	FLOAT|	TRUE
|CHAMBER 1|	UINT8	|1200|	100|	1|	BYTE|	FALSE
|CHAMBER 1|	INT16	|1300|	100|	1|	16BITS|	TRUE
|CHAMBER 1|	UINT16|	1400|	100|	1|	16BITS|	FALSE
|CHAMBER 1|	INT32|	1500|	100|	1|	32BITS|	TRUE
|CHAMBER 1|	UINT32|	1700|	100|	1|	32BITS|	FALSE


_Constraint: all columns must contain values. A check is made at runtime to ensure this contraint is met.__

In this example, tags for CHAMBER 1 include FLOATS, UINT8, etc. A check during tag generation verifies that there are no overlaps in memory for the group/format. 


Column descriptions:

|COLUMN|DESCRIPTIION              |
|------|--------------------------|
|GROUP| Group name for tag types
|FORMAT|  tag data type BOOL, UINT8, UINT16,INT17, UINT32, INT32, FLOAT
|START_ADDRESS|MODBUS start address for the given data type and group
| LENGTH | Amound of space resersed for the given data type
|TS_FORMAT| TwinSoft format. This could be used for other sytems|
|TS_SIGNED| Twinsoft sign flag TRUE or FALSE



## tag_importer.py

Options:
 * --excel TEXT   Excel file containing tags and memory map  [required]
 * --xmlin TEXT   Exported tag XML file from Twinsoft  [required]
 * --xmlout TEXT  Output file of generated XML file  [required]
 * --verbose      Will print more messages on console
 *  --help         Show this message and exit.

Commands:
*  clone     Clone folder from twinsoft export XML file Most cases Tags are...
*  generate  Generate tags using pattern defined in XL
*  tabulate  Tabulate input data and copy results to clipboard 

generate command:

Options
 * pattern - regex reflecting  which tags under TAG_PATTERN defined in excel TAGS tab  are to be generated [required]

clone command:

Options

* --tag_filter TEXT       Twinsoft Tag Name filter regex pattern Default: .+
* --group_filter TEXT     Twinsoft Group  regex pattern  [required]
* --dest TEXT             Destination Folder in Twinsoft. If not provided,
                          mirror group_filter pattern

* --loop TEXT             Loop number to ensure tags ang groups are unique
                          [required]

* --offset INTEGER        Address Offset to shift tags into  [required]
* --replace_pattern TEXT  Replacement filter regex pattern. Default: \d
* --help                  Show this message and exit.


tabulate command:

Arguments:  
* item: xmlsummary | tags | map | template

xmlsummary - summarizes modbus adressing by group and format
e.g.

 ||Group|   Format|  Signed| MB_MIN|  MB_MAX|
 |---|---|---|---|---|---|
|0|                  CHAMBER 1|  DIGITAL|   False|   1030|    1057
|1|                  CHAMBER 1|    FLOAT|    True|   1000|    1024
|2|                 CHAMBER 10|  DIGITAL|   False|   7030|    7057
|3|                 CHAMBER 10|    FLOAT|    True|   7000|    7024
|4|          CHAMBER 10\LOCALS|   16BITS|   False|   7700|    7837
|5|          CHAMBER 10\LOCALS|  DIGITAL|   False|   7500|    7548
|6|          CHAMBER 10\LOCALS|    FLOAT|    True|   7600|    7784
|7|           CHAMBER 1\LOCALS   |16BITS|   False|   1700|    1837

* tags - displays the contents of the TAB sheet in the specified Excel file
* map - validates and displays the contents of the MEMORY_MAP sheet in the specified Excel file
* template - validates and displays the contents of the TEMPLATE sheet in the specified Excel file

## Exceptions
|ERR CODE| DESCRIPTION|
|---|----|
|EE_TAB_NOT_FOUND = -200| tab/sheet not found in excel file
|EE_FILE_NOT_FOUND = -201| excel file not found
|EE_TAB_EMPTY = -203 | tab/sheet empty in excel file
|EE_EMPTY_CELLS = -204| tab/sheet contains in excel file has empty cells
|TE_PATTERN_NOT_FOUND = -100|regex pattern did not filter out any tags
|TE_XML_NOT_FOUND = -102|TwinSoft XML file not found
|TE_XML_ROOT_KEY_NOT_FOUND = -103|rootkey in TwinSoft XML not found. Mostly like not a valid export file.
|TE_XML_ATTRIBUTE_KEY_NOT_FOUND = -104|attribute key in TwinSoft XML not found. Mostly like not a valid export file.
|TE_GROUP_NOT_FOUND = -105|TAG entry has no corresponding GROUP in MEMORY_MAP
|TE_TAGS_EXIST = -107|Generated tag already exists in XML export file
|TE_TAG_NAME_TOO_LONG = -108|Generated tag long is too long (max 15 chars)
|TE_TAG_DESC_TOO_LONG = -109|Generated tag description too long (max 50 chars)
|TE_DUPLICATE_BOOL_ADDR = -110|Generated modbus address for BOOL type overlaps with other types in the MEMORY_MAP
|TE_DUPLICATE_ANALOG_ADDR = -111|Generated modbus address for ANALOG type overlaps with other types in the MEMORY_MAP
|TE_DUPLICATE_TAG_NAME = -112|
|TE_MEMORY_MAP_CONFLICT = -113|Duplicate TAG_NAME or TAG_PATTERN in TAG sheet in excel file

