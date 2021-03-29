# targetのモニター番号を引数で受け取る
# 指定がなければ1

if($args.Length -eq 0){
    $target = 1;
} else {
    $target = $args[0]
}

# モニター情報取得
Add-Type -AssemblyName System.Windows.Forms
$screens = [System.Windows.Forms.Screen]::AllScreens

# 指定された番号が存在するモニターの数より多ければ1番として扱う
if($target -gt $screens.count){
    $target = 1
}

# 配列が0始まりなので調整
$target = $target -1

# カーソルの位置調整
$width_center = $screens[$target].WorkingArea.X + ($screens[$target].WorkingArea.Width / 2)
$height_center = $screens[$target].WorkingArea.Y + ($screens[$target].WorkingArea.Height / 2)

# カーソル描画
# なぜか1回だけだと1→3の時だけwidthがおかしくなるので2回やる
[System.Windows.Forms.Cursor]::Position = new-object System.Drawing.Point($width_center, $height_center)
[System.Windows.Forms.Cursor]::Position = new-object System.Drawing.Point($width_center, $height_center)