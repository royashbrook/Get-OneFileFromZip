Import-Module .\Get-OneFileFromZip\Get-OneFileFromZip.psm1 -Force
Describe "Get-OneFileFromZip" {
    Context "When test.zip and test.txt" {
        BeforeAll{
            $x = Get-OneFileFromZip test.zip test.txt
        }
        It "should be String" {
            $x.GetType().Name | Should -Be 'String'
        }
        It "should have length of 5" {
            $x.length | Should -Be 5
        }
        It "should be 'test1'" {
            $x | Should -Be 'test1'
        }
    }
    Context "When test.zip and test.xml" {
        BeforeAll{
            [xml]$x = Get-OneFileFromZip test.zip test.xml
        }
        It "should be XmlDocument" {
            $x.GetType().Name | Should -Be 'XmlDocument'
        }
        It "a+b should be c when all cast int" {
            $a = [int]$x.root.a
            $b = [int]$x.root.b
            $c = [int]$x.root.c
            $a + $b | Should -Be $c
        }
        It "a+b should be 23 when string" {
            $a = $x.root.a
            $b = $x.root.b
            $a + $b | Should -Be '23'
        }
    }
    Context "When test.xlsx and sheet1.xlsx" {
        BeforeAll{
            [xml]$x = Get-OneFileFromZip test.xlsx sheet1.xml
        }
        It "should be XmlDocument" {
            $x.GetType().Name | Should -Be 'XmlDocument'
        }
        It "sheet should have one row, cell, and value and it should be 1" {
            $x.worksheet.sheetData.row.c.v | Should -Be "1"
        }
    }
}
Describe "Get-SheetByName($Path,$SheetName)" {
    BeforeAll{
        [xml]$sheet1 = Get-SheetByName test.xlsx sheet1
        [xml]$sheet2 = Get-SheetByName test.xlsx sheet2
        [xml]$sheet3 = Get-SheetByName test.xlsx sheet3
        [xml]$sheet4 = Get-SheetByName test.xlsx sheet4
        [xml]$sheet5 = Get-SheetByName test.xlsx sheet5
        [xml]$sheet6 = Get-SheetByName test.xlsx sheet6
        [xml]$sheet7 = Get-SheetByName test.xlsx sheet7
        [int]$sheet1a1 = ($sheet1.worksheet.sheetData.row.c | Where-Object r -eq "a1").v
        [int]$sheet2a1 = ($sheet2.worksheet.sheetData.row.c | Where-Object r -eq "a1").v
        [int]$sheet3a1 = ($sheet3.worksheet.sheetData.row.c | Where-Object r -eq "a1").v
        [int]$sheet4a1 = ($sheet4.worksheet.sheetData.row.c | Where-Object r -eq "a1").v
        [int]$sheet5a1 = ($sheet5.worksheet.sheetData.row.c | Where-Object r -eq "a1").v
        [int]$sheet6a1 = ($sheet6.worksheet.sheetData.row.c | Where-Object r -eq "a1").v
        [int]$sheet7a1 = ($sheet7.worksheet.sheetData.row.c | Where-Object r -eq "a1").v
        [int]$sumall = ($sheet1a1,$sheet2a1,$sheet3a1,$sheet4a1,$sheet5a1,$sheet6a1 | Measure-Object -Sum).Sum
    }
    Context "When test.xlsx and sheet1" {
        It "A1 should be 1" { $sheet1a1 | Should -Be "1" }
    }
    Context "When test.xlsx and sheet2" {
        It "A1 should be 2" { $sheet2a1 | Should -Be "2" }
    }
    Context "When test.xlsx and sheet3" {
        It "A1 should be 3" { $sheet3a1 | Should -Be "3" }
    }
    Context "When test.xlsx and sheet4" {
        It "Sheet4!A1 should be equal to Sheet1!A1" { $sheet4a1 | Should -Be $sheet1a1 }
    }
    Context "When test.xlsx and sheet5" {
        It "Sheet5!A1 should be equal to Sheet4!A1" { $sheet5a1 | Should -Be $sheet4a1 }
    }
    Context "When test.xlsx and sheet6" {
        It "Sheet6!A1 should be equal to Sheet3!A1" { $sheet6a1 | Should -Be $sheet3a1 }
    }
    Context "When test.xlsx and sheet7" {
        It "Sheet7!A1 should be equal to Sum Sheets 1-6" { $sheet7a1 | Should -Be $sumall }
    }
}