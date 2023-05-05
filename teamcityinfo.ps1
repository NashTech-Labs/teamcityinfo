param(
    [Parameter(Mandatory = $true)]
    [string] $TeamcityToken,
    [Parameter(Mandatory = $true)]
    [string] $Server_Url
)

Get-Module ImportExcel -ListAvailable | Import-Module -Force -Verbose

function get_headers {
    $headers = @{
        Authorization="Bearer $TeamcityToken"
    }
    return $headers
}

function get_project_info {
    [CmdletBinding()]
    param (
        [string] $url,
        [string] $suburl 
    )
    $uri = $url + $suburl
    Write-Host $uri
    $headers = get_headers
    try {
        $data =Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
    }
    catch {
        $data = $_.ErrorDetails
    }
    return $data    
}

$projects_xml_data = get_project_info -url $Server_Url -suburl "/app/rest/projects"
$projects = $projects_xml_data.projects.project

$project = @{
    project_name=""
    project_id=""
    project_parentId=""
    project_webUrl=""
    latest_build_id=""
    latest_build_number=""
    latest_build_status=""
    latest_build_url=""
    latest_build_finish_date=""
    latest_build_finish_time=""
}

$project_count = 0
foreach ($project_data in $projects) {
    Write-Host "--------------- Processing FOR $($project_data.name)-------"
    $project["project_name"] = $project_data.name
    $project["project_id"] = $project_data.id
    $project["project_webUrl"] = $project_data.webUrl
    if($project_data.id -eq "_Root"){
        Write-Host "--------------SKIPPED Root--------------------"
        continue
    }
    $project["project_parentId"] = $project_data.parentProjectId
    $latest_build_url = "/app/rest/builds?locator=project:$($project_data.id),running:any,count:1"
    $latest_build_xml_data = get_project_info -url $Server_Url -suburl $latest_build_url
    if(([string]$latest_build_xml_data).Contains("404")){
        $project["latest_build_id"]="NA"
        $project["latest_build_number"] = "NA"
        $project["latest_build_status"] = "NA"
        $project["latest_build_finish_time"] = "NA"
    }
    else{
        $latest_build_data = $latest_build_xml_data.builds
        $count = $latest_build_data.count
        if($count -eq "0"){
            $project["latest_build_id"]="NA"
            $project["latest_build_number"] = "NA"
            $project["latest_build_status"] = "NA"
            $project["latest_build_url"] = "NA"
            $project["latest_build_finish_time"] = "NA"
            $project["latest_build_finish_date"] = "NA"
        }
        else{
            $project["latest_build_id"]= $latest_build_data.build.id
            $project["latest_build_number"] = $latest_build_data.build.number
            $project["latest_build_url"] = $latest_build_data.build.webUrl
            $latest_build_state = $latest_build_data.build.state
            if($latest_build_state -eq "running"){
                $project["latest_build_finish_date"] = Get-Date
                $project["latest_build_finish_date"] = $project["latest_build_finish_date"].ToString("yyyy/MM/dd")
                $project["latest_build_finish_time"] = Get-Date
                $project["latest_build_finish_time"] = $project["latest_build_finish_time"].ToString("HH:mm:ss:ff:zzzz")
                $project["latest_build_status"] = "RUNNING"
            }
            else{
                $project["latest_build_status"] = $latest_build_data.build.status
                $project["latest_build_finish_date"] = [datetime]::ParseExact($latest_build_data.build.finishOnAgentDate,"yyyyMMddTHHmmsszzz",[Globalization.CultureInfo]::CurrentCulture).ToString("yyyy/MM/dd") 
                $project["latest_build_finish_time"] = [datetime]::ParseExact($latest_build_data.build.finishOnAgentDate,"yyyyMMddTHHmmsszzz",[Globalization.CultureInfo]::CurrentCulture).ToString("HH:mm:ss:ff:zzzz")
            }
        }
    }
    
    $build_id = $project["latest_build_id"]
    if( $build_id -eq "NA"){
        $project["agent_name"]= "NA"
        $project["agent_id"] = "NA"
        $project["latest_build_start_date"]="NA"
        $project["VCS"] = "NA"
        $project["RepositoryUrl"] = "NA"
        $project["ArtifactoryUrl"] = "NA"
    }
    else {
        $suburl = "/app/rest/builds/id:$build_id"
        $build_data = get_project_info -url $Server_Url -suburl $suburl

        $project["agent_name"]= $build_data.build.agent.name
        $project["latest_build_start_date"] = [datetime]::ParseExact($build_data.build.startDate,"yyyyMMddTHHmmsszzz",[Globalization.CultureInfo]::CurrentCulture).ToString("yyyy/MM/dd") 
        if($build_data.build.state -eq "finished"){
            $project["latest_build_finish_date"] = [datetime]::ParseExact($build_data.build.finishDate,"yyyyMMddTHHmmsszzz",[Globalization.CultureInfo]::CurrentCulture).ToString("yyyy/MM/dd") 
        }
        else{
            $project["latest_build_finish_date"] = "NA"
        }

        $project["ArtifactoryUrl"] = "NA"

        if($build_data.build.properties){
            if($build_data.build.properties.property){
                foreach ($property_data in $build_data.properties.property) {
                    if ($property_data.name -eq "ArtifactoryUrl") {
                        $project["ArtifactoryUrl"] = $property_data.value
                    }
                }
            }
        }            
    }

    $vcs_git_suburl = "/app/rest/vcs-roots?locator=project:$($project["project_id"]),count:1" #type:jetbrains.git"
    $vcs_data = get_project_info -url $Server_Url -suburl $vcs_git_suburl
    
    if(($vcs_data.'vcs-roots'.count -eq 0) -or (([string]$vcs_data).Contains("404"))){
        $project["RepositoryUrl"] = "NA"
        $project["VCS"] = "NA"
    }
    else {
        $suburl = $vcs_data.'vcs-roots'.'vcs-root'.href
        $vcs_git_data = get_project_info -url $Server_Url -suburl $suburl
        if(([string]$vcs_data).Contains("404")){
            $project["RepositoryUrl"] = "NA"
            $project["VCS"] = "NA"
        }
        else{
            $project["VCS"] = $vcs_git_data.'vcs-root'.vcsName
            if($vcs_git_data.'vcs-root'.properties){
                if($vcs_git_data.'vcs-root'.properties.property){
                    foreach($property_data in $vcs_git_data.'vcs-root'.properties.property){
                        if($property_data.name -eq "url"){
                            $project["RepositoryUrl"] = $property_data.value.Replace("%GitHub.URL%","https://github.com/duck-creek")
                            break
                        }
                    }
                }
            }
        }
    }
    
    $project_count+=1
    $project["Serial No."] = $project_count
    $project
    # Sending data to json file
    $jsonfile = "$PSScriptRoot\project_data.json"
    $projects_data_json = Get-Content $jsonfile | Out-String | ConvertFrom-Json
    $projects_file_data = [System.Collections.ArrayList] $projects_data_json
    $projects_file_data.Add($project) | out-null
    $projects_file_data | ConvertTo-Json | Set-Content $jsonfile    
    
    Write-Host "--------------- DONE FOR $($project_data.name) and total projects done are $project_count -------------"
}

Get-Content "$PSScriptRoot\project_data.json" | ConvertFrom-Json | Export-Excel -Path "$PSScriptRoot\project_data.xlsx"