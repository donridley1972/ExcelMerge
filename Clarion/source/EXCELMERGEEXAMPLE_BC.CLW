  MEMBER('ExcelMergeExample.clw')

  PRAGMA('define(init_priority=>3)')

  MAP
    MODULE('ExcelMergeExample_BC0.CLW')
ExcelMergeExample_BC0:DctInit             PROCEDURE()
ExcelMergeExample_BC0:DctKill             PROCEDURE()
ExcelMergeExample_BC0:FilesInit           PROCEDURE()
    END
  END

DctInit PROCEDURE()
  CODE
  ExcelMergeExample_BC0:DctInit
  ExcelMergeExample_BC0:FilesInit


DctKill PROCEDURE()
  CODE
  ExcelMergeExample_BC0:DctKill

