
function Export-Excel {

    <#   
    .SYNOPSIS  
        Export data boject to excel file.
        
    .DESCRIPTION  
        Export obect to an excel file.
        
    .PARAMETER InputObject
        Object to add to excel document.
    
    .PARAMETER Path
        Path and filename of the excel file to write.
        
    .EXAMPLE  
    Get-ChildItem *.csv | Export-Excel -Path 'report.xlsx'
    
    #>
     
    # Requires -version 2.0  
    [CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'low', DefaultParameterSetName = 'file' )]

    Param (
        [Parameter(ValueFromPipeline = $True, Mandatory = $True, HelpMessage = "Object to import.")] [ValidateNotNullOrEmpty()] [System.Object[]] $InputObject,
        [Parameter(ValueFromPipeline = $False, Mandatory = $True, HelpMessage = "Name of excel file output")] [ValidateNotNullOrEmpty()] [string] $Path    
    )

    BEGIN {     
        $header = @();
        $rowIndex = 1;

        $excel = New-Object -ComObject excel.application;
        $excel.DisplayAlerts = $False;
        $excel.Visible = $False;

        $workbook = $excel.WorkBooks.add(1);
        $worksheet = $workbook.WorkSheets.item(1);
    }

    PROCESS {
        foreach ( $object in $InputObject ) {
            ## Build header row if first iteration.
            if ($rowIndex -eq 1) {
                $i = 0;
                $object | Get-Member -MemberType NoteProperty | ForEach-Object {
                    $worksheet.cells.item(1, ++$i) = $_.Name;
                    $header += $_.Name;
                };
                $rowIndex++;
            }
        
            ## Insert values into current row index.
            $colIndex = 1;            
            foreach ($name in $header) {
                $value = $object."$name";
                $worksheet.cells.item($rowIndex, $colIndex) = $value;
                $colIndex++;
            }
            $rowIndex++;
        }
    }        

    END {
        ## Save spreadsheet and clear reference.
        $workbook.saveas($Path);
        
        $workbook.Close($false);
        $excel.Quit();

        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        [void] [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet)
        [void] [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

        Remove-Variable -Name excel

    }
}
