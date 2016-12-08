###############################
########## FUNCTIONS ##########
###############################




########################################
########## NAMING CONVENTIONS ##########
########################################

function loadNamingConventions($NamingConventions)
{
    # Read CSV with prefixes. The columns are:
    # Name;Prefix;Type;SSIS2012;SSIS2012VS2015;SSIS2014;SSIS2016
    $csv = Import-Csv $NamingConventions -delimiter ";"

    # Create hash table for storing the prefixes and lookup
    # values. A hash table is easier to search in
    $hashtablePrefix = @{}

    # Loop through CSV rows
    foreach($Row in $csv)
    {
        # The dirty try machanisme is to prevent duplicate keys
        # need to fix that
        try
        {
            if ($Row.SSIS2012.Trim() -ne "N/A")
            {
                $hashtablePrefix.Add($Row.SSIS2012.Trim(),$Row.Prefix.Trim())
            }
        } catch {}
        try
        {
            if ($Row.SSIS2012VS2015.Trim() -ne "N/A")
            {
                $hashtablePrefix.Add($Row.SSIS2012VS2015.Trim(),$Row.Prefix.Trim())
            }
        } catch {}
        try
        {
            if ($Row.SSIS2014.Trim() -ne "N/A")
            {
                $hashtablePrefix.Add($Row.SSIS2014.Trim(),$Row.Prefix.Trim())
            }
        } catch {}
        try
        {
            if ($Row.SSIS2016.Trim() -ne "N/A")
            {
                $hashtablePrefix.Add($Row.SSIS2016.Trim(),$Row.Prefix.Trim())
            }
        } catch {}

    }
    return $hashtablePrefix
}


########################################
########## LOG IN TABLE ##########
########################################

function addLogRow($Solution, $Project, $Package, $Path, $Name, $Prefix, $Error)
{
    #Create a row
    $row = $errorTable.NewRow()

    #Enter data in the row
    $row.Solution = $Solution 
    $row.Project = $Project
    $row.Package = $Package
    $row.Path = $Path
    $row.Name = $Name
    $row.Prefix = $Prefix
    $row.Error = $Error

    #Add the row to the table
    $errorTable.Rows.Add($row)
} 

function checkNamingConvention($SolutionName, $ProjectName, $PackageName, $Path, $Name, $Key, $Flow)
{
    
    $Prefix = FindPrefix($Key)
    #write-host $Path "|" $Name "|" $Key "|" $Flow "|" $Prefix
    if ($Prefix -eq "")
    {
        $message = "Unknown " + $Flow + " Flow Item:" + $Key
        addLogRow $SolutionName $ProjectName $PackageName $Path $Name "?" $message
    }
    elseif (-not $Name.StartsWith($Prefix))
    {
        $message = "Wrong Prefix " + $Flow + " Flow Item"
        addLogRow $SolutionName $ProjectName $PackageName $Path $Name $Prefix $message
    }
    #else
    #{
    #    #Do Nothing,
    #}

}


function loopTasksAndComponents($TasksContainers, $PackagePath)
{
   $PackageName = Split-Path $PackagePath -Leaf
   $ProjectName = Split-Path (Split-Path $PackagePath -Parent) -Leaf
   $SolutionName = Split-Path (Split-Path (Split-Path $PackagePath -Parent) -Parent) -Leaf

   $Prefix = ""
   #$TotalBad = 0
   #$TotalGood = 0

    foreach ($TaskContainer in $TasksContainers)
    {
    
        if ($TaskContainer.Executables)
        {
            loopTasksAndComponents $TaskContainer.Executables.Executable $PackagePath
        }
 
        # Tasks
        #write-host $TaskContainer.RefId "|" $TaskContainer.ObjectName "|" $TaskContainer.ExecutableType
        checkNamingConvention $SolutionName $ProjectName $PackageName $TaskContainer.RefId $TaskContainer.ObjectName $TaskContainer.ExecutableType "Control"
        

        # Check
        #$Prefix = FindPrefix($TaskContainer.ExecutableType)
        #if ($Prefix -eq "")
        #{
        #    addLogRow $SolutionName $ProjectName $PackageName $TaskContainer.RefId $TaskContainer.ObjectName "?" "Unknown Control Flow Item 0:" $TaskContainer.ExecutableType
        #}
        #elseif (-not $TaskContainer.ObjectName.StartsWith($Prefix))
        #{
        #    addLogRow $SolutionName $ProjectName $PackageName $TaskContainer.RefId $TaskContainer.ObjectName $Prefix "Wrong Prefix Control Flow Item"
        #    #$TotalBad = $TotalBad + 1
        #}
        #else
        #{
        #    #$TotalGood = $TotalGood + 1
        #}

         
        # Components within DataFlowTask
        if ($TaskContainer.ExecutableType -eq "Microsoft.Pipeline" -or                  # SSIS 2012 
            $TaskContainer.ExecutableType -eq "SSIS.Pipeline.3" -or                     # SSIS 2014 and later
            $TaskContainer.ExecutableType -eq "{5918251B-2970-45A4-AB5F-01C3C588FE5A}") # SSIS 2012 via DFT Tab
        {
            foreach ($comp in $TaskContainer.ObjectData.pipeline.components.component)
            {
                # Handle components with same ComponentClassID
                if ($comp.componentClassID -eq "Microsoft.ManagedComponentHost" -or     # SSIS 2012
                    $comp.componentClassID -eq "DTS.ManagedComponentWrapper.3" -or      # SSIS 2012 in VS2015
                    $comp.componentClassID.StartsWith("{"))                             # SSIS 2014 and later
                {
                    $propfound = $false
                    foreach ($prop in $comp.properties.property)
                    {
                        # Search for property UserComponentTypeName
                        if ($prop.name -eq "UserComponentTypeName")
                        {
                            #write-host  $comp.refId "|" $comp.name "|" $prop.innertext
                            checkNamingConvention $SolutionName $ProjectName $PackageName $comp.RefId $comp.name $prop.innertext "Data"

                            # Check
                            #$Prefix = FindPrefix($prop.innertext)
                            #if ($Prefix -eq "")
                            #{
                            #    addLogRow $SolutionName $ProjectName $PackageName $comp.refId $comp.name "?" "Unknown Data Flow Item 1:" $prop.innertext
                            #}
                            #elseif (-not $comp.name.StartsWith($Prefix))
                            #{
                            #    addLogRow $SolutionName $ProjectName $PackageName $comp.refId $comp.name $Prefix "Wrong Prefix Data Flow Item"
                            #    #$TotalBad = $TotalBad + 1
                            #}
                            #else
                            #{
                            #    #$TotalGood = $TotalGood + 1
                            #}

                            $propfound = $true
                        }
                    }
                    # If property UserComponentTypeName not found use componentClassID (GUID)
                    if ($propfound -eq $false)
                    {
                        #write-host  $comp.refId "|" $comp.name "|" $comp.componentClassID
                        checkNamingConvention $SolutionName $ProjectName $PackageName $comp.RefId $comp.name $comp.componentClassID "Data"

                        ## Check
                        #$Prefix = FindPrefix($comp.componentClassID)
                        #if ($Prefix -eq "")
                        #{
                        #    addLogRow $SolutionName $ProjectName $PackageName $comp.refId $comp.name "?" "Unknown Data Flow Item 2:" $comp.componentClassID
                        #}
                        #elseif (-not $comp.name.StartsWith($Prefix))
                        #{
                        #    addLogRow $SolutionName $ProjectName $PackageName $comp.refId $comp.name $Prefix "Wrong Prefix Data Flow Item"
                        #    #$TotalBad = $TotalBad + 1
                        #}
                        #else
                        #{
                        #    #$TotalGood = $TotalGood + 1
                        #}
                    }
                }
                # Handle regular components with unique componentClassID (no GUID)
                else
                {
                    #write-host  $comp.refId "|" $comp.name "|" $comp.componentClassID
                    checkNamingConvention $SolutionName $ProjectName $PackageName $comp.RefId $comp.name $comp.componentClassID "Data"

                    ## Check
                    #$Prefix = FindPrefix($comp.componentClassID)
                    #if ($Prefix -eq "")
                    #{
                    ##write-host  $comp.refId "|" $comp.name "|" $comp.componentClassID
                    #    addLogRow $SolutionName $ProjectName $PackageName $comp.refId $comp.name "?" "Unknown Data Flow Item 3:" $comp.componentClassID
                    #}
                    #elseif (-not $comp.name.StartsWith($Prefix))
                    #{
                    #    addLogRow $SolutionName $ProjectName $PackageName $comp.refId $comp.name $Prefix "Wrong Prefix Data Flow Item"
                    #    #$TotalBad = $TotalBad + 1
                    #}
                    #else
                    #{
                    #    #$TotalGood = $TotalGood + 1
                    #}
                }
            }
        }
    }

}


function FindPrefix($search)
{
    $returnvalue = ""
    Foreach ($Key in ($hashtablePrefix.GetEnumerator() | Where-Object {$_.Name -eq $search}))
    {
        $returnvalue = $Key.Value
    }
    return $returnvalue
}