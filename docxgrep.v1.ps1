# beta1 writes everything to a single csv or xlsx file

# Main configuration variables
#$path = "C:\Users\ohm\Box\Faculty Exam Submission\Paul Keyword Code\Test\"

$path = "C:\Users\zl516\Box\1_WPDOCS\1 - Exams & Grades Team\Exams\Exam Administrator\Original Exams\24A Exams\Final Exams\In-Class Exams\Faculty Exam Submissions\"
#$path = "C:\Users\zl516\Box\1_WPDOCS\1 - Exams & Grades Team\Exams\Exam Administrator\Original Exams\24A Exams\Final Exams\Take-home Exams\Faculty Exam Submissions\"

$stop_words_file = "$($path)\00keywords.txt"
$characters_around = 150
$output_file = "$($path)\00keywordhits.csv"

# Just docx xfiles
$files = Get-ChildItem -Path "$($path)*" -Include *.docx,*.doc

# Initialize variables
$results = @()
$filehits = 0

# For now, will do something bad if there is an empty line at the end of the stop words.
# Fix might involve trim somehow
# The \b at the start means keywords must match at the start of a word
# The lack of a \b at the end means keywords will match even with varying suffixes

$findtext = "\b(" + ((Get-Content $stop_words_file) -join "|") + ")"

echo "Searching for keywords in all .docx and .doc files in $path"

Function getStringMatch
{
	# Start MS Word
	$application = New-Object -comobject Word.Application

	Foreach ($file in $files)
	{
		echo $file
		$doc = $application.documents.open($file.FullName, $false, $true)
		$range = $doc.content
		$contents = $range.Text
		$con_len = $contents.Length

		#Foreach ($hit in $hits.Matches)
		$matches = ([regex]$findtext).Matches($contents)

		if ($matches.Length -gt 0)
		{
			$filehits++
			echo "    Hit!!!"
		}
		else
		{
			# No matches, so populate a single empty entry
			$properties = @{
				    File = $file.Name
				    Match = "NO HITS"
				    TextAround = "NO HITS FOUND in file $($file.Name)"
				    Loc = -1
				    Percent = "{0:P}" -f 0
			}
			$results += @(New-Object -TypeName PsCustomObject -Property $properties)
		}

		Foreach($match in $matches)
		{
			#echo $match
			$hit_ind = $match.index
			$left = $hit_ind - $characters_around

			# Edge case: context would push left before left boundary
			if ($left -lt 0) {
			   $left = 0
			}

			# Edge case: context would push length after end of file
			$snippet_length = $match.length + ($characters_around * 2)
			if (($left + $snippet_length) -gt $con_len)
			{
				$snippet_length = $con_len - $left
			}

			$context = $contents.Substring($left, $snippet_length)
			
			$properties = @{
				    File = $file.Name
				    Match = $match
				    TextAround = $context
				    Loc = $hit_ind
				    Percent = "{0:P}" -f ($hit_ind / $con_len)
			}
			
			$results += @(New-Object -TypeName PsCustomObject -Property $properties)
		 }
		 $doc.close()

		 # Add an empty row to the csv
		 $properties = @{
			    Match = ""
		}
		$results += @(New-Object -TypeName PsCustomObject -Property $properties)

	}

	if ($results)
	{
		$results | Select-Object -Property File,Loc,Percent,Match,TextAround | Export-Csv $output_file -NoTypeInformation
	}

	# Reporting
	echo ""
	echo "Final Results"
	echo ""
	echo "$($files.Length) processed"
	echo "$filehits files had hits"
	echo "$($results.Length) total hits"
	
	$application.quit()
}

# Measure-Command times how long it takes to run
Measure-Command { getStringMatch | Out-Default }

