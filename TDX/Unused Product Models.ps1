. (Join-Path $PSSCRIPTROOT "TDX-API.ps1")

# Get all TDX Assets
$allTDXAssets = Search-TDXAssets -AppName ITAsset

# Get all TDX Product Models
$allProductModels = Get-TDXAssetProductModels

# Set up an empty array for unused models
$unusedModels = New-Object System.Collections.Generic.List[System.Object]

# going through each model
foreach ($model in $allProductModels) {
    [int]$pct = ($allProductModels.IndexOf($model) / $allProductModels.count) * 100
    Write-progress -Activity "Working on $($model.Name)" -percentcomplete $pct -status "$pct% Complete"

    # based on the model name, find assets that are active or on hand
    $matchedProductModel = $allTDXAssets | Where-Object { ($_.ProductModelName -eq $model.Name) -and ($_.StatusName -eq 'Active' -or $_.StatusName -eq 'On Hand') }

    # if there are no assets, save to unused array
    if ($matchedProductModel.count -eq 0) {
        $unusedModels.Add($model.Name)

        # Setup the api call options
        $productModelOptions = @{
            ID             = $model.ID 
            Name           = $model.Name 
            ManufacturerID = $model.ManufacturerID 
            ProductTypeID  = $model.ProductTypeID  
            Description    = $model.Description 
            PartNumber     = $model.PartNumber 
            IsActive       = $false
        }

        # Modify the Product Model to inactive
        Edit-TDXAssetProductModel @productModelOptions
    }
}
Write-progress -Activity 'Done...' -Completed