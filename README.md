# TestStatisticsExport

Export plugin with teststatistics for Tests.
Compatible with ILIAS 4.4, 5.0 and 5.1.

###  creates a xlsx-file with teststatistics (Item-total correlation, difficulty, distractoranalysis)###
 
 See [thread at ilias.de](http://www.ilias.de/docu/goto_docu_frm_2528_3388.html)
 
### Usage ###
Install the plugin

```bash
mkdir -p Customizing/global/plugins/Modules/Test/Export  
cd Customizing/global/plugins/Modules/Test/Export
git clone https://github.com/kyro46/TestStatisticsExport.git
```

and activate it in the ILIAS-Admin-GUI. Activate manual correction.

You now have an additional export option for tests.

### Credits ###
* Development for ILIAS 4.4+ by Christoph Jobst, University Leipzig 2015/2016
 
 This Project uses PHPExcel. See [PHPExcel GitHub](https://github.com/PHPOffice/PHPExcel) and the [PHPExcel Homepage](https://phpexcel.codeplex.com/), License: GNU LGPL
 
