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