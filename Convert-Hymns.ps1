# About
# 
# 1. There are more eliquent ways of doing this, however my goal was ease and not perfection :-) (For example, I could have done both Psalms and Hymns in one script)
# 2. The script depends on detecting new Psalms using the keyword "Psalm" followed by space and digits.
# 3. The script depends on detecting new Hymns using the keyword "Hymn" followed by space and digits.
# 4. The script depends on detecting new verses by using paragraphs beginning with digits.  As a result, songs with 1 verse (Psalm 127, Hymn 1 etc..) will have empty content.  You can avoid this by adding a "1. " before these verses.
# 5. Copyright information is ignored in this script.  It is your responsibility to add copyright information for Hymns 38, 50, 66, 79. You will need to have CCLI and CCLI streaming licenses for these if applicable.
# 6. The script is dependent on proper formating of the word document.  It is possible that it isn't 100% accurate.  Still need to keep an eye out, but should be 99% accurate

# Pre-Requisites/Instructions
# 1. Licensed Word Document purchased from Premier Printing at bookofpraise.ca
# 2. Copy just the Psalms into a new docx file (psalms.docx)
# 3. Copy just the Hymns into a new docx file (hymns.docx)
# 4. Find and replace all manual page breaks "^m" with Paragraph marks "^p"

# Settings
$HymnsFile = "C:\Users\mdede\Documents\Church\Sound\Slideshows\Hymns.docx"
$pptxTemplate = "C:\Users\mdede\Documents\Church\Sound\Slideshows\BOPtemplate.pptx"
$pptxOutputDir = "C:\temp\HymnsTest" #Note: Script threw errors for me if this path had spaces.

#Open Word Doc
$wordapp = New-Object -ComObject word.application
$HymnsDocx = $wordapp.Documents.Open($HymnsFile)
$wordapp.visible = 'msoTrue'
$Paragraphs = $HymnsDocx.content.paragraphs

#Add Powerpoint types
Add-type -AssemblyName office -ErrorAction SilentlyContinue
Add-Type -AssemblyName microsoft.office.interop.powerpoint -ErrorAction SilentlyContinue

#For Each Paragraph in Word doc
$errors = @()
ForEach($i in 1..($Paragraphs).count){
    
    #If Paragraph text matches text "Hymn" followed by space and 1 or 2 digits
    If($Paragraphs[$i].Range.Text -match "Hymn \d{1,2}"){
        #If we are currently adding slides to an existing presentation (we need to close and save the presentation before continuing)
        if($pres){
            Try{
                #The first slide is our template slide.  Delete it.
                $firstslide = $pres.Slides | select -first 1
                $firstslide.Delete()

                #Save the slide
                $pres.SaveAs($filename,[Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsDefault)
                if($filename -notin $errors){$copyPres = $pres.SaveCopyAs($filename)}
                
                #Close the presentation
                $pres.close()

                #If we don't close Powerpoint here, it will eventually run out of memory and our com objects will fail
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject( $pres )
                Stop-Process -Name 'Powerpnt' -force -ErrorAction SilentlyContinue
                $pres = $null
            }Catch [Exception]{
                #do nothing
            }
        } #End saving previous presentation

        #We have identified a new song.  Set the title and filename info.
        $title = (Get-Culture).TextInfo.ToTitleCase($Paragraphs[$i].range.text.ToLower())
        $Filename = ($Paragraphs[$i].range.text + '.pptx') -replace "\s+",""
        $filename = "$pptxOutputDir\$filename"

        #Test if we already have a file with the expected filename.  Otherwise continue.
        if(Test-path $filename){
            #already processed
        }else{
            #Output new filename
            $filename

            #Open template Powerpoint presentation
            $PPTXapp = New-Object -ComObject powerpoint.application
            $pres = $PPTXapp.Presentations.open($pptxTemplate)
            start-sleep 3
            $PPTXapp.visible = "msoTrue"
            $slides = $pres.slides

        }

    #If paragraph begins with a digit and we don't have a file saved and we haven't encountered an error for this new file. (It is a new verse for the song we are currently processing)
    }elseif($Paragraphs[$i].range.text -match "^\d{1,2}" -and !(Test-path $filename) -and !($filename -in $errors)){
        
        Try{
            #In the case where we completed a previous verse for the song, $nextSlide is the previous verse, otherwise it is null. Find the shape that has the keyword '<Verse>' and remove it.  
            ForEach($shape in $nextSlide.shapes){
                if($shape.TextFrame.TextRange.Text.IndexOf('<Verse>') -ge 0){
                    $shape.TextFrame.TextRange.Text = $shape.TextFrame.TextRange.Text -replace '<Verse>',''
                    $nextslide = $null
                }
            }
        }Catch [Exception]{
            #expected error.  We can't enumerate slides that don't exist
        }

        #Take the first slide and create a copy
        Try{
            $firstslide = $slides | select -first 1
            $firstSlide.copy()
            start-sleep -milliseconds 500
            $nextslide = $slides.paste()
        }catch [Exception]{
            $errors += $filename
        }
        start-sleep -Milliseconds 100
        
        #Keep enumerating the Word Document until we find last paragraph with text
        $j = $i
        Do{
            $j++
        }while($paragraphs[$j].Range.Text -notmatch "^\d{1,2}" -and $Paragraphs[$j].Range.Text -notmatch '^\s+\Z')
        $endIndex = $j - 1

        #Select and copy the entire verse we identified.
        $selectedRange = $wordapp.ActiveDocument.Range($wordapp.ActiveDocument.Paragraphs($i).Range.Start,$wordapp.ActiveDocument.Paragraphs($endIndex).Range.End)
        $selectedRange.copy()
        
        #Replace the keywords in the slides
        ForEach($k in 1..$nextslide.shapes.count){
            Try{
                #Replace <title>
                if($nextslide.shapes[$k].TextFrame.TextRange.Text -match '<Title>'){ 
                    start-sleep -Milliseconds 100
                    $nextslide.shapes[$k].TextFrame.TextRange.Text = ($nextslide.shapes[$k].TextFrame.TextRange.Text -replace '<Title>',$title)
                }
                #<CopyrightInfo> is replaced with null...this was too hard to get to work neatly
                if($nextslide.shapes[$k].TextFrame.TextRange.Text -match '<CopyrightInfo>' -ge 0){ 
                    start-sleep -Milliseconds 100
                    $nextslide.shapes[$k].TextFrame.TextRange.Text = $nextslide.shapes[$k].TextFrame.TextRange.Text -replace '<CopyrightInfo>',''
                }
            }Catch [Exception]{
                $errors += $filename
            }
            start-sleep -Milliseconds 100
            if($nextslide.shapes[$k].TextFrame.TextRange.Text -match '<Verse>'){
                Try{
                    #Paste verse text
                    $nextslide.shapes[$k].TextFrame.TextRange.paste() | out-null
                   
                    #What color do you want font?
                    $nextslide.shapes[$k].TextFrame.TextRange.Font.Color.rgb = [System.Drawing.Color]::white.ToArgb()
                    
                    #Do you want it bold?
                    $nextslide.shapes[$k].TextFrame.TextRange.Font.Bold = 0
                    
                    #What size do you want the font?
                    $nextslide.shapes[$k].TextFrame.TextRange.Font.Size = 32

                    #What font name?
                    $nextslide.shapes[$k].TextFrame.TextRange.Font.NameAscii = 'Calabri'

                    #Do you want bullet points?
                    $nextslide.shapes[$k].TextFrame.TextRange.ParagraphFormat.Bullet = 0

                    #How much space do you want before a paragraph?
                    $nextslide.shapes[$k].TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0

                    #How much space do you want after a paragraph?
                    $nextslide.shapes[$k].TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0

                    #What line spacing?
                    $nextslide.shapes[$k].TextFrame.TextRange.ParagraphFormat.Spacewithin = 1
                    
                }Catch [Exception]{
                    $errors += $filename
                }
            }
        }

    }
}

