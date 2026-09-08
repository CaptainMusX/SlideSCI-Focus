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
$group=$ribbon.Tabs[0].Groups | Where-Object {$_.Name -eq '图片处理'}
$script:checks=0
function Check($ok,$label) { if(-not $ok){throw $label}; $script:checks++; "PASS $label" }
Check ($group.Items.Count -eq 5) 'two action columns, two separators and one format column'
$left=$group.Items[0]; $right=$group.Items[2]
Check ($left.Items.Count -eq 3 -and $right.Items.Count -eq 3) 'identical action column structure'
Check ($left.Items[2].SizeString -eq $right.Items[2].SizeString) 'identical offset widths'
Check ($left.Items[0].ControlSize.ToString() -eq $right.Items[0].ControlSize.ToString() -and $left.Items[1].ControlSize.ToString() -eq $right.Items[1].ControlSize.ToString()) 'matching button control sizes'
for($i=0;$i -lt 2;$i++) {
 Check ([Object]::ReferenceEquals($left.Items[$i].Image,$right.Items[$i].Image)) 'same icon instance in matching rows'
 Check ($left.Items[$i].Label.Length -eq $right.Items[$i].Label.Length) 'same label lengths in matching rows'
}
$format=$group.Items[4]
Check ($format.BoxStyle.ToString() -eq 'Vertical' -and $format.Items.Count -eq 3) 'third column is one vertical container'
$title=$format.Items[0]; $font=$format.Items[1]; $rowSlot=$format.Items[2]; $row=$rowSlot.Items[0]
Check ($title.GetType() -eq $font.GetType() -and $title.SizeString -eq $font.SizeString) 'same native type and size for wide inputs'
Check ($title.Label.Length -eq $font.Label.Length) 'wide input labels align'
Check ($row.Items.Count -eq 3 -and $row.BoxStyle.ToString() -eq 'Horizontal') 'single horizontal third row'
Check ($row.Items[0].SizeString -eq $left.Items[2].SizeString) 'font size matches both offsets'
Check ($row.Items[2].Items.Count -eq 4 -and @($row.Items[2].Items | Where-Object {$_.GetType().Name -match 'Toggle'}).Count -eq 0) 'four ordinary alignment buttons'
# Serialize the actual compiled Ribbon with the VSTO runtime.
$writerType=$impl.GetType('Microsoft.Office.Tools.Ribbon.RibbonManagerImpl+RibbonFactory')
$writer=[Activator]::CreateInstance($writerType,$flags,$null,@('Microsoft.PowerPoint.Presentation',$false,$ribbon),$null)
[xml]$xml=$writerType.GetProperty('RibbonXml',$flags).GetValue($writer,$null)
$ns=[Xml.XmlNamespaceManager]::new($xml.NameTable); $ns.AddNamespace('r',$xml.DocumentElement.NamespaceURI)
$actual=$xml.SelectSingleNode('//r:group[@id="图片处理"]',$ns)
Check ($actual.SelectNodes('./r:comboBox',$ns).Count -eq 0 -and $actual.SelectNodes('./r:box',$ns).Count -eq 3) 'serialized group keeps three top-level columns'
[IO.File]::WriteAllText((Join-Path $output 'vsto-ribbon.xml'),$xml.OuterXml)
if($CreatePreview) {
 $controls=@{}
 function Collect($item) { $controls[$item.Id]=$item; if($item.PSObject.Properties['Items']) { foreach($child in $item.Items){Collect $child} } }
 Collect $group
 Collect ($ribbon.Tabs[0].Groups | Where-Object {$_.Name -eq '图片自动对齐'})
 $tabs=$xml.SelectSingleNode('//r:tabs',$ns)
 $tab=$xml.CreateElement('tab',$xml.DocumentElement.NamespaceURI); $tab.SetAttribute('id','TitleLayoutPreview'); $tab.SetAttribute('label','标题布局验收')
 $null=$tab.AppendChild($xml.SelectSingleNode('//r:group[@id="图片自动对齐"]',$ns).CloneNode($true))
 $null=$tab.AppendChild($actual.CloneNode($true)); $tabs.RemoveAll(); $null=$tabs.AppendChild($tab)
 $images=@{}; $imageIndex=0
 foreach($element in @($xml.SelectNodes('//*'))) {
  $control=$controls[$element.GetAttribute('id')]
  foreach($attribute in @($element.Attributes)) { if($attribute.Name -match '^(get|on)' -or $attribute.Name -in @('loadImage','tag')){$element.RemoveAttribute($attribute.Name)} }
  if(-not $control){continue}
  foreach($property in @('Label','ShowLabel','ShowImage','Enabled','Visible')) {
   if($control.PSObject.Properties[$property]) {
    $value=$control.$property
    if($null -ne $value -and "$value" -ne '') { $element.SetAttribute($property.Substring(0,1).ToLower()+$property.Substring(1),$(if($value -is [bool]){$value.ToString().ToLower()}else{"$value"})) }
   }
  }
  if($control.PSObject.Properties['ControlSize']){$element.SetAttribute('size',$(if($control.ControlSize.ToString() -match 'Large'){'large'}else{'normal'}))}
  if($control.PSObject.Properties['OfficeImageId'] -and $control.OfficeImageId){$element.SetAttribute('imageMso',$control.OfficeImageId)}
  elseif($control.PSObject.Properties['Image'] -and $control.Image){
   $imageIndex++; $key="rIdImage$imageIndex"; $path=Join-Path $output "$key.png"
   $control.Image.Save($path,[Drawing.Imaging.ImageFormat]::Png); $images[$key]=$path; $element.SetAttribute('image',$key)
  }
 }
 # Static preview has no callbacks or macros and does not alter add-in registration.
 $app=New-Object -ComObject PowerPoint.Application; $deck=$app.Presentations.Add(0)
 $preview=Join-Path $output ('TitleRibbon-'+[Guid]::NewGuid().ToString('N').Substring(0,6)+'.pptx')
 try{$null=$deck.Slides.Add(1,12);$deck.SaveAs($preview,24)}finally{$deck.Close()}
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $zip=[IO.Compression.ZipFile]::Open($preview,'Update')
 try {
  function ZipText($name,$text){$entry=$zip.CreateEntry($name);$stream=[IO.StreamWriter]::new($entry.Open());try{$stream.Write($text)}finally{$stream.Dispose()}}
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
