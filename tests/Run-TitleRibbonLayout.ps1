<# Source/serialized-Ribbon regression only. -CreatePreview creates a separate
   static PPTX for manual inspection; these checks do not prove pixel alignment. #>
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
Check ($ribbon.Tabs[0].Visible -and $ribbon.Tabs[0].ControlId.CustomId -eq 'CaptainMusX_SlideSCIFocus_SciFigure') 'SciFigure has a product-specific visible tab identity'
Check ($ribbon.Tabs[1].ControlId.CustomId -eq 'CaptainMusX_SlideSCIFocus_SciStudio') 'SciStudio has a different product-specific identity'
$align=$ribbon.Tabs[0].Groups | Where-Object {$_.Name -eq '图片自动对齐'}
Check (@($align.Items | Where-Object {$_.SizeString -eq '0000'}).Count -eq 5) 'auto-arrangement input widths preserved'
$group=$ribbon.Tabs[0].Groups | Where-Object {$_.Name -eq '图片处理'}
Check ($group.Items.Count -eq 5) 'title group has three columns and two separators'
$left=$group.Items[0]; $right=$group.Items[2]; $format=$group.Items[4]
foreach($column in @($left,$right,$format)) {
 Check ($column.Items.Count -eq 3) 'title column has three rows'
 foreach($row in $column.Items) { Check ($row.BoxStyle.ToString() -eq 'Horizontal') 'all title rows have equal container depth' }
}
Check ($left.Items[2].Items[0].Label -eq '垂直偏移' -and $right.Items[2].Items[0].Label -eq '水平偏移') 'offset controls preserved'
Check ($format.Items[2].Items.Count -eq 3) 'third title row contains exactly three controls'
$size=$format.Items[2].Items[0]; $alignment=$format.Items[2].Items[1]; $grouping=$format.Items[2].Items[2]
Check ($size.Label -eq '字号' -and $size.GetType().Name -match 'ComboBox') 'complete editable font-size combo remains first'
Check ($alignment.Name -eq 'titleAlignmentMenu' -and $alignment.Items.Count -eq 4) 'alignment menu remains second'
Check ($grouping.Label -eq '编组' -and $grouping.GetType().Name -match 'Toggle') 'grouping toggle remains third'
Check ($format.Items[0].Items[0].SizeString -eq $format.Items[1].Items[0].SizeString) 'title and font have matching input widths'
Check ($size.Items.Count -eq 29 -and $size.Items[0].Label -eq '2' -and $size.Items[28].Label -eq '200') 'font-size presets preserved'
$populate=$dll.GetType('SlideSCI.Ribbon1').GetMethod('PopulateTitleFontSizePresets',$flags)
$arguments=New-Object object[] 1; $arguments[0]=[string[]]@('8','12','24')
$originalPresets=[string[]]@($size.Items | ForEach-Object {$_.Label})
$size.Text='13.5';$null=$populate.Invoke($ribbon,$arguments);$null=$populate.Invoke($ribbon,$arguments)
Check ($size.Items.Count -eq 3 -and $size.Text -eq '13.5') 'rebuilding presets preserves typed size and avoids duplicates'
$arguments[0]=$originalPresets;$null=$populate.Invoke($ribbon,$arguments)
$settingsType=$dll.GetType('SlideSCI.Properties.Settings')
$settingsDefault=$settingsType.GetProperty('Default',[Reflection.BindingFlags]'Static,Public,NonPublic').GetValue($null,$null)
$historyProperty=$settingsType.GetProperty('TitleTextHistory',$flags)
$originalHistory=$historyProperty.GetValue($settingsDefault)
$refreshHistory=$dll.GetType('SlideSCI.Ribbon1').GetMethod('RefreshTitleHistoryCombo',$flags)
try {
 $historyProperty.SetValue($settingsDefault, (@('A','A','B','C','D','E','F') -join [char]10))
 $null=$refreshHistory.Invoke($ribbon,$null)
 Check ((@($format.Items[0].Items[0].Items | ForEach-Object {$_.Label}) -join ',') -eq 'A,B,C,D,E') 'history deduplicates and limits entries to five'
} finally { $historyProperty.SetValue($settingsDefault,$originalHistory); $null=$refreshHistory.Invoke($ribbon,$null) }
$zoom=$ribbon.Tabs[0].Groups | Where-Object {$_.Name -eq 'zoomGroup'}
Check ($zoom.Items.Count -eq 3) 'zoom keeps three columns'
$col1=$zoom.Items[0];$col2=$zoom.Items[1];$col3=$zoom.Items[2]
foreach($column in $zoom.Items) { Check ($column.Items.Count -eq 3) 'zoom column keeps three rows' }
Check ($col1.Items[0].Items[0].Label -eq '框选区域' -and $col2.Items[0].Items[0].Label -eq '放大选区') 'zoom action labels updated'
Check ($col1.Items[0].Items[0].ShowImage -and $col2.Items[0].Items[0].ShowImage) 'zoom actions retain icons'
Check ($col1.Items[1].Items[0].SizeString -eq $col2.Items[1].Items[0].SizeString) 'zoom size fields use identical widths'
Check ($col1.Items[1].Items[0].SizeString -eq '00%') 'size width reserves space for visible units'
Check ((@($col2.Items[1].Items[0].Items | ForEach-Object {$_.Label}) -join ',') -eq '1x,2x,3x,4x,5x,10x') 'only requested magnification presets are offered'
Check (@($col1.Items[1].Items[0].Items | Where-Object {$_.Label -notlike '*%'}).Count -eq 0) 'percentage presets display their unit'
foreach($column in @($col1,$col2)) {
 Check ($column.Items[2].Items.Count -eq 2 -and $column.Items[2].Items[0].GetType().Name -match 'Label') 'stroke label is independent of the menu'
 $menu=$column.Items[2].Items[1]
 Check ($menu.GetType().Name -match 'RibbonMenu') 'stroke selector is a menu, not a split button'
 Check (@($menu.Items | Where-Object {$_.GetType().Name -match 'Gallery'}).Count -ge 2) 'theme and standard color galleries exist'
 Check (@($menu.Items | Where-Object {$_.Label -eq '无轮廓'}).Count -eq 1) 'no-outline option exists'
 Check (@($menu.Items | Where-Object {$_.Label -eq '箭头'}).Count -eq 1) 'arrow options exist'
 Check (@($menu.Items | Where-Object {$_.Label -eq '取色器'}).Count -eq 1) 'eyedropper entry exists'
 $populateStroke=$dll.GetType('SlideSCI.Ribbon1').GetMethod('PopulateStrokeMenu',$flags)
 $before=$menu.Items.Count
 $null=$populateStroke.Invoke($ribbon,@($menu,[Object]::ReferenceEquals($column,$col1)))
 Check ($menu.Items.Count -eq $before) 'reopening dynamic outline menu does not duplicate controls'
}
$writerType=$impl.GetType('Microsoft.Office.Tools.Ribbon.RibbonManagerImpl+RibbonFactory')
$writer=[Activator]::CreateInstance($writerType,$flags,$null,@('Microsoft.PowerPoint.Presentation',$false,$ribbon),$null)
[xml]$xml=$writerType.GetProperty('RibbonXml',$flags).GetValue($writer,$null)
$ns=[Xml.XmlNamespaceManager]::new($xml.NameTable);$ns.AddNamespace('r',$xml.DocumentElement.NamespaceURI)
$titleXml=$xml.SelectSingleNode('//r:group[@id="图片处理"]',$ns)
Check ($titleXml.SelectNodes('./r:box',$ns).Count -eq 3) 'serialized title group has three columns'
Check ($titleXml.SelectSingleNode('./r:box[3]/r:box[3]',$ns).ChildNodes.Count -eq 3) 'serialized font-size, alignment and grouping share the third row'
Check (@($xml.SelectNodes('//*[@id]') | Group-Object id | Where-Object Count -gt 1).Count -eq 0) 'serialized control IDs are unique'
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
