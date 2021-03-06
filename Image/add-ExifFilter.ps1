Function Add-exifFilter {
<#
        .Synopsis
            Adds an Exif Filter to a list of filters, or creates a new filter
        .Description
            Adds an Exif Filter to a list of filters, or creates a new filter
        .Example
            Add-exifFilter -passThru -ExifID $ExifIDKeywords -typeid 1101 -string "Ocean,Bahamas"    |      
            Adds a filter to set the keywords to Ocean; Bahamas, using the numeric type ID 
            and getting the function to convert the string to the vector type required
         .Example
            Add-exifFilter -passThru -ExifID $ExifIDTitle -typeName "vectorofbyte" -string "fish"
            Add a filter to set the Title to "fish", using the name of the type 
            and getting the function to convert the string to the vector type required
         .Example   
            Add-exifFilter -passThru -ExifID $ExifidCopyright -typeName "String" -value "© James O'Neill 2009" 
            Sets the copyright field (note this is a normal string, not a vector of bytes containing a unicode string)
         .Example      
            Add-exifFilter -passThru -ExifID $ExifIDGPSAltitude -typeName "uRational" -Numerator 123 -denominator 10
            Add a filter to set the GPS Altitude to 12.3M 
            getting the function to create the unsigned rational required
        .Parameter ExifID
            The ID of the field to be added / updated
        .Parameter TypeID
            The code representing the data type for this field (String, byte, integer, ratio, vector etc)
        .Parameter Value
            The new value for the field
        .Parameter TypeName
            Reserved Will allow the type to specified as a name rather than a numeric code
        .Parameter Numerator
            Reserved will ratios to be passed as numerator / denominator
        .Parameter Denominator
            Reserved will ratios to be passed as numerator / denominator
        .Parameter String
            Reserved will allow the value for Vectors which hold strings to be passed as a string
        .Parameter passthru
            If set, the filter will be returned through the pipeline.  This should be set unless the filter is saved to a variable.
        .Parameter filter
            The filter chain that the rotate filter will be added to.  If no chain exists, then the filter will be created
    #>

param(
    [Parameter(ValueFromPipeline=$true)]
    [__ComObject]
    $filter,
               
    [Parameter(Mandatory=$true)]$ExifID, 
    $typeid , $value , $string , $Numerator, $denominator , $typeName,
   
    [switch]$passThru                      
    )
    
    process {
        if (-not $filter) { $filter = New-Object -ComObject Wia.ImageProcess } 
        if ($typeName -and -not $typeiD) {$typeid = @{"Undefined"=1000; "Byte"=1001;"String"=1002;"uInt"=1003;"Long"=1004;"uLong"=1005;"Rational"=1006;"URational"=1007
                                                      "VectorOfUndefined"=1100; "VectorOfByte"=1101; "VectorOfUint" = 1102; "VectorOfLong"= 1103; "VectorOfULong"= 1104; "VectorOfRational" = 1105; "VectorOfURational" = 1106;}[$typeName] }
        if ((-not $filter.Apply) -or (-not $typeID)) { return }
        
        if ((@(1006,1007) -contains $Typeid) -and (-not $value) -and ($numerator -ne $null) -and $denominator) {
            $value =New-Object -ComObject wia.rational                                                                                                                                                                                                         
            $value.Denominator = $denominator                                                                                                                                                                                                                     
            $value.Numerator = $Numerator                                                                                                                                                                                                                      
        }
        if ((@(1100,1101) -contains $TypeID) -and (-not $value) -and $string) {$value = New-Object -ComObject "WIA.Vector"
                                                                                 $value.SetFromString($string)
        }
        if ((1002 -eq $TypeID) -and (-not $value) -and $string) {$value = $string } 
        
        
        $filter.Filters.Add($filter.FilterInfos.Item("Exif").FilterId)
        $filter.Filters.Item($filter.Filters.Count).Properties.Item("ID")   = "$ExifID"       
        $filter.Filters.Item($filter.Filters.Count).Properties.Item("Type") = "$TypeID"
        $filter.Filters.Item($filter.Filters.Count).Properties.Item("Value")= $Value 
        if ($passthru) { return $filter }         
    }
}

$ExifUndefined                 = 1000
$ExifByte                      = 1001
$ExifString                    = 1002
$ExifUnsignedInteger           = 1003
$ExifLong                      = 1004
$ExifUnsignedLong              = 1005
$ExifRational                  = 1006
$ExifUnsignedRational          = 1007
$ExifVectorOfUndefined         = 1100
$ExifVectorOfBytes             = 1101
$ExifVectorOfUnsignedIntegers  = 1102
$ExifVectorOfLongs             = 1103
$ExifVectorOfUnsignedLongs     = 1104
$ExifVectorOfRationals         = 1105
$ExifVectorOfUnsignedRationals = 1106
