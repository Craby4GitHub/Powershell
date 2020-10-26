function Center-Control {
    param (
        $ParentLocation,
        $ParentSize,
        $ChildLocation,
        $ChildSize
    )
    $ChildLocation.Location.X = $ParentLocation.Location.X + (($ParentSize.Width / 2) - ($ChildSize.Width / 2))
    $ChildLocation.Location.Y = $ParentLocation.Location.Y - 20
    return $ChildLocation
}