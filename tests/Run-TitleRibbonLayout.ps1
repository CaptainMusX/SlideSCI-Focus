<#
  SciFigure 功能区布局回归：图片自动排列输入宽度、添加图片标题四列布局、
  标题历史下拉、局部放大三列功能分组与标签对齐。用 -CreatePreview 生成可打开的 PPTX 静态预览，
  用于真实 PowerPoint 渲染验收（不做自动截图）。
#>
[CmdletBinding()]
param([switch]$CreatePreview)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$output=Join-Path $root 'artifacts/title-ribbon-preview'
$null=New-Item -ItemType Directory -Force $output
$impl=[Reflection.Assembly]::LoadFrom('C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.Office.Tools.Common.Implementation\v4.0_10.0.0.0__b03f5f7f11d50a3a\Microsoft.Office.Tools.Common.Implementation.dll')
$flags=[Reflection.BindingFlags]'Instance,Public,NonPublic'
$factory=[Activator]::CreateInstance($impl.GetType('Microsoft.Office.Tools.Ribbon.RibbonFactoryImpl'),$flags,$null,@($null),$null)
$dll=[Reflection.Assembly]::LoadFrom((Join-Path $root 'SlideSCI/bin/Release/CaptainMusX.SlideSCI.Focus.dll'))
$ctor=$dll.GetType('SlideSCI.Ribbon1').GetConstructors($flags) | Where-Object {$_.GetParameters().Length -eq 1}
$ribbon=$ctor.Invoke(@($factory))
$script:checks=0
function Check($ok,$label) { if(-not $ok){throw $label}; $script:checks++; "PASS $label" }

# ---- 图片自动排列：五个输入框宽度缩为 2/3（sizeString 6 位 -> 4 位）----
$align=$ribbon.Tabs[0].Groups | Where-Object {$_.Name -eq '图片自动对齐'}
$narrow=@($align.Items | Where-Object {$_.SizeString -eq '0000'})
Check ($narrow.Count -eq 5) 'auto-arrange inputs use the narrowed 2/3 size string'
Check (@($align.Items | Where-Object {$_.SizeString -eq '000000'}).Count -eq 0) 'no auto-arrange input keeps the old wide size string'

# ---- 添加图片标题：四列，第三列三行 ComboBox，第四列编组与对齐 ----
$group=$ribbon.Tabs[0].Groups | Where-Object {$_.Name -eq '图片处理'}
Check ($group.Items.Count -eq 7) 'title group is four columns with three separators'
$left=$group.Items[0]; $right=$group.Items[2]; $format=$group.Items[4]; $options=$group.Items[6]
Check ($left.Items.Count -eq 3 -and $right.Items.Count -eq 3) 'identical action column structure'
Check ($left.Items[2].Label -eq '垂直偏移' -and $right.Items[2].Label -eq '水平偏移') 'offset labels are vertical/horizontal'
Check ($left.Items[2].SizeString -eq $right.Items[2].SizeString) 'identical offset widths'
Check ($left.Items[0].ControlSize.ToString() -eq $right.Items[0].ControlSize.ToString() -and $left.Items[1].ControlSize.ToString() -eq $right.Items[1].ControlSize.ToString()) 'matching button control sizes'
for($i=0;$i -lt 2;$i++) {
 Check ([Object]::ReferenceEquals($left.Items[$i].Image,$right.Items[$i].Image)) 'same icon instance in matching rows'
 Check ($left.Items[$i].Label.Length -eq $right.Items[$i].Label.Length) 'same label lengths in matching rows'
}
Check ($format.BoxStyle.ToString() -eq 'Vertical' -and $format.Items.Count -eq 3) 'third column holds three rows'
Check ($format.Items[0].Label -eq '标题' -and $format.Items[1].Label -eq '字体' -and $format.Items[2].Label -eq '字号') 'third column rows are title, font and font size'
foreach($item in $format.Items) { Check ($item.GetType().Name -match 'ComboBox') 'third column rows are all combo boxes' }
Check ($format.Items[0].SizeString -eq $format.Items[1].SizeString) 'title and font share the wide width'
Check ($format.Items[2].SizeString -eq $left.Items[2].SizeString) 'font size matches both offsets'
Check ($options.BoxStyle.ToString() -eq 'Vertical' -and $options.Items.Count -eq 2) 'fourth column holds grouping and alignment'
Check ($options.Items[0].GetType().Name -match 'Toggle' -and $options.Items[0].Label -eq '编组') 'grouping toggle in the fourth column'
Check ($options.Items[1].Items.Count -eq 4 -and @($options.Items[1].Items | Where-Object {$_.GetType().Name -match 'Toggle'}).Count -eq 0) 'four alignment buttons in the fourth column'

# ---- 字号 ComboBox：内置预设，输入框与下拉箭头一体 ----
Check ($format.Items[2].Items.Count -eq 29 -and $format.Items[2].Items[0].Label -eq '2' -and $format.Items[2].Items[28].Label -eq '200') 'font size presets are ready before Ribbon Load'
$populate=$dll.GetType('SlideSCI.Ribbon1').GetMethod('PopulateTitleFontSizePresets',$flags)
$arguments=New-Object object[] 1; $arguments[0]=[string[]]@('8','12','24')
$originalSize=$format.Items[2].Text
$originalPresets=[string[]]@($format.Items[2].Items | ForEach-Object {$_.Label})
$format.Items[2].Text='13.5'
$null=$populate.Invoke($ribbon,$arguments)
Check ((@($format.Items[2].Items | ForEach-Object {$_.Label}) -join ',') -eq '8,12,24') 'font size preset values and order preserved'
$null=$populate.Invoke($ribbon,$arguments)
Check ($format.Items[2].Items.Count -eq 3 -and $format.Items[2].Text -eq '13.5') 'rebuilding presets neither duplicates entries nor changes typed size'
$arguments[0]=$originalPresets; $null=$populate.Invoke($ribbon,$arguments)
$format.Items[2].Text=$originalSize

# ---- 标题历史：最新 5 条、去重、超限淘汰 ----
$settingsType=$dll.GetType('SlideSCI.Properties.Settings')
$staticFlags=[Reflection.BindingFlags]'Static,Public,NonPublic'
$settingsDefault=$settingsType.GetProperty('Default',$staticFlags).GetValue($null,$null)
$historyProperty=$settingsType.GetProperty('TitleTextHistory',$flags)
$originalHistory=$historyProperty.GetValue($settingsDefault)
$refreshHistory=$dll.GetType('SlideSCI.Ribbon1').GetMethod('RefreshTitleHistoryCombo',$flags)
$historyProperty.SetValue($settingsDefault, (@('A','B','C','D','E','F') -join [char]10))
$null=$refreshHistory.Invoke($ribbon,$null)
Check ((@($format.Items[0].Items | ForEach-Object {$_.Label}) -join ',') -eq 'A,B,C,D,E') 'title history keeps the latest five entries'
$historyProperty.SetValue($settingsDefault, (@('A','A','B') -join [char]10))
$null=$refreshHistory.Invoke($ribbon,$null)
Check ((@($format.Items[0].Items | ForEach-Object {$_.Label}) -join ',') -eq 'A,B') 'title history removes duplicate entries'
$historyProperty.SetValue($settingsDefault,$originalHistory)
$null=$refreshHistory.Invoke($ribbon,$null)

# ---- 局部放大：三列功能分组，每列三行单控件 ----
$zoom=$ribbon.Tabs[0].Groups | Where-Object {$_.Name -eq 'zoomGroup'}
Check ($zoom.Items.Count -eq 3) 'zoom group keeps three columns'
$col1=$zoom.Items[0]; $col2=$zoom.Items[1]; $col3=$zoom.Items[2]
Check ($col1.BoxStyle.ToString() -eq 'Vertical' -and $col1.Items.Count -eq 3) 'selection column has three rows'
Check ($col1.Items[0].Label -eq '选区' -and $col1.Items[1].Label.Trim() -eq '尺寸' -and $col1.Items[2].Label.TrimStart([char]0x200A) -match '^框线') 'selection column is 选区/尺寸/框线'
Check ($col2.Items[0].Label -eq '放大' -and $col2.Items[1].Label.Trim() -eq '尺寸' -and $col2.Items[2].Label -match '^引线') 'generate column is 放大/尺寸/引线'
Check ($col3.Items[0].Label.Trim() -eq '间距' -and $col3.Items[1].Label -match '^连线' -and $col3.Items[2].Label -eq '编组') 'options column is 间距/连线/编组'
Check ($col3.Items[2].GetType().Name -match 'Toggle') 'zoom grouping control is a toggle button'
foreach($col in @($col1,$col2,$col3)) { foreach($item in $col.Items) { Check ($item.GetType().Name -ne 'RibbonBox') 'zoom columns have no nested layout box' } }
Check ($col1.Items[1].SizeString -eq $col3.Items[0].SizeString) 'selection size and gap share the same narrow width'

$nbsp=[string][char]0x00A0
$fine=[string][char]0x200A
foreach($combo in @($col1.Items[1],$col2.Items[1],$col3.Items[0])) {
 Check ($combo.Label.StartsWith($nbsp)) 'zoom input label has a fixed optical inset'
 Check (-not $combo.Text.Contains($nbsp) -and -not $combo.Text.Contains($fine)) 'optical inset never enters numeric text'
}
Check ($col1.Items[1].Label.StartsWith($nbsp+$fine) -and $col2.Items[1].Label.StartsWith($nbsp+$fine)) 'first two zoom columns use the fine inset'
Check (-not $col3.Items[0].Label.StartsWith($nbsp+$fine)) 'third zoom column keeps the base inset'

# ---- 序列化真实 Ribbon XML ----
# 确认第三列是三个 ComboBox，第四列是编组/对齐，且不存在横向盒。
$writerType=$impl.GetType('Microsoft.Office.Tools.Ribbon.RibbonManagerImpl+RibbonFactory')
$writer=[Activator]::CreateInstance($writerType,$flags,$null,@('Microsoft.PowerPoint.Presentation',$false,$ribbon),$null)
[xml]$xml=$writerType.GetProperty('RibbonXml',$flags).GetValue($writer,$null)
$ns=[Xml.XmlNamespaceManager]::new($xml.NameTable); $ns.AddNamespace('r',$xml.DocumentElement.NamespaceURI)
$titleXml=$xml.SelectSingleNode('//r:group[@id="图片处理"]',$ns)
Check ($titleXml.SelectNodes('./r:box',$ns).Count -eq 4) 'serialized title group keeps four top-level columns'
Check ($titleXml.SelectNodes('.//r:box[@boxStyle="horizontal"]',$ns).Count -eq 0) 'title group has no horizontal formatting box'
Check ($titleXml.SelectNodes('./r:box[3]/*',$ns).Count -eq 3 -and $titleXml.SelectNodes('./r:box[3]/r:comboBox',$ns).Count -eq 3) 'serialized third column is three combo boxes'
Check ($titleXml.SelectNodes('./r:box[4]/*',$ns).Count -eq 2) 'serialized fourth column is grouping and alignment'
$zoomXml=$xml.SelectSingleNode('//r:group[@id="zoomGroup"]',$ns)
Check ($zoomXml.SelectNodes('./r:box',$ns).Count -eq 3) 'serialized zoom group keeps three top-level columns'
Check ($zoomXml.SelectNodes('.//r:box[@boxStyle="horizontal"]',$ns).Count -eq 0) 'zoom group has no horizontal box that shifts a row upward'
[IO.File]::WriteAllText((Join-Path $output 'vsto-ribbon.xml'),$xml.OuterXml)

if($CreatePreview) {
 $controls=@{}
 function Collect($item) { $controls[$item.Id]=$item; if($item.PSObject.Properties['Items']) { foreach($child in $item.Items){Collect $child} } }
 Collect $group
 Collect ($ribbon.Tabs[0].Groups | Where-Object {$_.Name -eq '图片自动对齐'})
 Collect ($ribbon.Tabs[0].Groups | Where-Object {$_.Name -eq 'zoomGroup'})
 $tabs=$xml.SelectSingleNode('//r:tabs',$ns)
 $tab=$xml.CreateElement('tab',$xml.DocumentElement.NamespaceURI); $tab.SetAttribute('id','LayoutPreview'); $tab.SetAttribute('label','布局验收')
 $null=$tab.AppendChild($xml.SelectSingleNode('//r:group[@id="图片自动对齐"]',$ns).CloneNode($true))
 $null=$tab.AppendChild($xml.SelectSingleNode('//r:group[@id="zoomGroup"]',$ns).CloneNode($true))
 $null=$tab.AppendChild($titleXml.CloneNode($true))
 $tabs.RemoveAll(); $null=$tabs.AppendChild($tab)
 $images=@{}; $imageIndex=0
 foreach($element in @($xml.SelectNodes('//*'))) {
  $control=$controls[$element.GetAttribute('id')]
  foreach($attribute in @($element.Attributes)) { if($attribute.Name -match '^(get|on)' -or $attribute.Name -in @('loadImage','tag')){$element.RemoveAttribute($attribute.Name)} }
  if(-not $control){continue}
  foreach($property in @('Label','ShowLabel','ShowImage','Enabled','Visible')) {
   if($property -eq 'Enabled' -and $element.LocalName -in @('box','separator','group','tab')) { continue }
   if($control.PSObject.Properties[$property]) {
    $value=$control.$property
    if($null -ne $value -and "$value" -ne '') { $element.SetAttribute($property.Substring(0,1).ToLower()+$property.Substring(1),$(if($value -is [bool]){$value.ToString().ToLower()}else{"$value"})) }
   }
  }
  if($control.PSObject.Properties['ControlSize'] -and -not $element.SelectSingleNode('ancestor::r:menu | ancestor::r:buttonGroup',$ns)){$element.SetAttribute('size',$(if($control.ControlSize.ToString() -match 'Large'){'large'}else{'normal'}))}
  if($control.PSObject.Properties['OfficeImageId'] -and $control.OfficeImageId){$element.SetAttribute('imageMso',$control.OfficeImageId)}
  elseif($control.PSObject.Properties['Image'] -and $control.Image){
   $imageIndex++; $key="rIdImage$imageIndex"; $path=Join-Path $output "$key.png"
   $control.Image.Save($path,[Drawing.Imaging.ImageFormat]::Png); $images[$key]=$path; $element.SetAttribute('image',$key)
  }
 }
 # Static preview has no callbacks or macros and does not alter add-in registration.
 $app=New-Object -ComObject PowerPoint.Application; $deck=$app.Presentations.Add(0)
 $preview=Join-Path $output ('LayoutPreview-'+[Guid]::NewGuid().ToString('N').Substring(0,6)+'.pptx')
 try{$null=$deck.Slides.Add(1,12);$deck.SaveAs($preview,24)}finally{$deck.Close()}
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $zip=[IO.Compression.ZipFile]::Open($preview,'Update')
 try {
  function ZipText($name,$text){$entry=$zip.CreateEntry($name);$stream=[IO.StreamWriter]::new($entry.Open(),[Text.UTF8Encoding]::new($false));try{$stream.Write($text)}finally{$stream.Dispose()}}
  function ReadZipXml($name){$entry=$zip.GetEntry($name);$reader=[IO.StreamReader]::new($entry.Open());try{[xml]$result=$reader.ReadToEnd()}finally{$reader.Dispose()};$entry.Delete();return $result}
  ZipText 'customUI/customUI.xml' $xml.OuterXml
  $rels=ReadZipXml '_rels/.rels'
  $rel=$rels.CreateElement('Relationship',$rels.DocumentElement.NamespaceURI);$rel.SetAttribute('Id','rIdRibbonPreview');$rel.SetAttribute('Type','http://schemas.microsoft.com/office/2006/relationships/ui/extensibility');$rel.SetAttribute('Target','customUI/customUI.xml');$null=$rels.DocumentElement.AppendChild($rel)
  ZipText '_rels/.rels' $rels.OuterXml
  [xml]$imageRels='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
  foreach($key in $images.Keys){
   $null=[IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip,$images[$key],"customUI/$key.png")
   $rel=$imageRels.CreateElement('Relationship',$imageRels.DocumentElement.NamespaceURI);$rel.SetAttribute('Id',$key);$rel.SetAttribute('Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');$rel.SetAttribute('Target',"$key.png");$null=$imageRels.DocumentElement.AppendChild($rel)
  }
  ZipText 'customUI/_rels/customUI.xml.rels' $imageRels.OuterXml
  $types=ReadZipXml '[Content_Types].xml'
  if(-not @($types.DocumentElement.ChildNodes | Where-Object {$_.Extension -eq 'png'}).Count){$png=$types.CreateElement('Default',$types.DocumentElement.NamespaceURI);$png.SetAttribute('Extension','png');$png.SetAttribute('ContentType','image/png');$null=$types.DocumentElement.AppendChild($png)}
  ZipText '[Content_Types].xml' $types.OuterXml
 } finally {$zip.Dispose()}
 "PREVIEW=$preview"
}
"TOTAL PASS $script:checks"
