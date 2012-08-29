#1
#
#install latest stable Ruby Version (1.9.2) - 28.08.12
#this version includes rubygems, you don't need to require them with this version
#
#gem install spreadsheet
#
#OR
#
#install latest version of rubygems (http://rubygems.org/pages/download)
#run setup.rb
#
#gem install spreadsheet
#
#you need to require rubygems
#
#2
#
#create folder "xls_to_csv" for each participant
#change "Spreadsheet.open('RAID202_onsets_14_08_2012.xls')"
########################################################################################

require 'rubygems'
require 'spreadsheet'

#initialize hashes for data you want to write to csv file
@control_c = {}
@control_e = {}
@drop_c = {}
@future_c = {}
@future_e = {}
@past_c = {}
@past_e = {}
@rating = {}
@future_detail = {}
@future_emotion = {}
@future_significance = {}
@future_valence = {}
@past_detail = {}
@past_emotion = {}
@past_significance = {}
@past_valence = {}

# generates a new xls file with number 420 for nil objects in condition into folder "output_xls_after_parsing_for_nil_objects"
# if you don't need this comment the next 15 lines (upto "Spreadsheet end")
Spreadsheet.open('RAID103_onsets_14_08_2012.xls') do |book| #open spreadsheet --> change name here for other spreadsheets
  book.worksheet('EPRI_ONSET_copied_values_only').each do |row| #select a specified worksheet from the above named spreadsheet
    break if row[80]  # iteration over rows breaks on 80  - maybe you have to change the number if number 80 is not your last "data row"
    count = 1  #counter for runs set to one for first run
    while count < 7 do #loop for six runs
	if row[2].nil? && row[0] == count #condition for nil objects in runs
		row[2] = 420 #write integer 420 to row condition if it's nil
		row[8...14] = [420, 420, 420, 420, 420, 420, 420] #write integer 420 to rows 8, 9, 10, 11, 12, 13, 14 if they're nil
		#print "new row #{row[8...14]}\n"
	end #if end
	count += 1 #add one for next run
     end #while end
  end #book end
  book.write 'output_xls_after_parsing_for_nil_objects/RAID103_onsets_14_08_2012.xls' #write changes to output file
end #Spreadsheet end

Spreadsheet.open('RAID103_onsets_14_08_2012.xls') do |book| #do you remember?
  book.worksheet('EPRI_ONSET_copied_values_only').each do |row| 
    break if row[80] 
    count = 1
    while count < 7 do	

	if row[2].nil? && row[0] == count
		row[8...14] = [420, 420, 420, 420, 420, 420, 420]
	end

	if row[2] == "control" && row[0] == count #condition for string "control" in column "condition"
		(@control_c[count] ||= []) << ["#{row[8]},"] #define new array "@control_c[count]" and push values for the specified run
	        (@control_e[count] ||= []) << ["#{row[9]},"]
         	file_control_c = File.new("xls_to_csv/control_c_#{count}.csv", "w+") # create new File object "file_control_c"
	 	#print "CONTROL_C #{@control_c[count]}\n" 
         		for e in @control_c[count] do # iterate over all index in array "@control_c[count]"
  	    			file_control_c.write e #write values to file object
         		end #end for
         	file_control_c.close #close file object
		file_control_e = File.new("xls_to_csv/control_e_#{count}.csv", "w+") #do you remember?
	 	print "CONTROL_E #{@control_e[count]}\n" 
         		for e in @control_e[count] do
  	    			file_control_e.write e
         		end #end for
         	file_control_e.close
	elsif row[2] == "drop" && row[0] == count
		(@drop_c[count] ||= []) << ["#{row[8]},"]
		file_drop_c = File.new("xls_to_csv/drop_c_#{count}.csv", "w+")
	 	print "DROP_C #{@drop_c[count]}\n" 
         		for e in @drop_c[count] do
  	    			file_drop_c.write e
         		end #end for
         	file_drop_c.close
	elsif row[2] == "future" && row[0] == count
		(@future_c[count] ||= []) << ["#{row[8]},"]
		(@future_e[count] ||= []) << ["#{row[9]},"]
		(@future_detail[count] ||= []) << ["#{row[11]},"]
		(@future_emotion[count] ||= []) << ["#{row[12]},"]
		(@future_valence[count] ||= []) << ["#{row[13]},"]
		(@future_significance[count] ||= []) << ["#{row[14]},"]
		file_future_c = File.new("xls_to_csv/future_c_#{count}.csv", "w+")
	 	print "FUTURE_C #{@future_c[count]}\n" 
         		for e in @future_c[count] do
  	    			file_future_c.write e
         		end #end for
         	file_future_c.close
		file_future_e = File.new("xls_to_csv/future_e_#{count}.csv", "w+")
	 	print "FUTURE_E #{@future_e[count]}\n" 
         		for e in @future_e[count] do
  	    			file_future_e.write e
         		end #end for
         	file_future_e.close
		file_future_detail = File.new("xls_to_csv/future_detail_#{count}.csv", "w+")
	 	print "FUTURE_DETAIL #{@future_detail[count]}\n" 
         		for e in @future_detail[count] do
  	    			file_future_detail.write e
         		end #end for
         	file_future_detail.close
		file_future_emotion = File.new("xls_to_csv/future_emotion _#{count}.csv", "w+")
	 	print "FUTURE_emotion  #{@future_emotion [count]}\n" 
         		for e in @future_emotion [count] do
  	    			file_future_emotion .write e
         		end #end for
         	file_future_emotion .close
		file_future_significance = File.new("xls_to_csv/future_significance _#{count}.csv", "w+")
	 	print "FUTURE_significance  #{@future_significance [count]}\n" 
         		for e in @future_significance [count] do
  	    			file_future_significance .write e
         		end #end for
         	file_future_significance .close
		file_future_valence = File.new("xls_to_csv/future_valence _#{count}.csv", "w+")
	 	print "FUTURE_valence  #{@future_valence [count]}\n" 
         		for e in @future_valence [count] do
  	    			file_future_valence .write e
         		end #end for
         	file_future_valence .close
	elsif row[2] == "past" && row[0] == count
		(@past_c[count] ||= []) << ["#{row[8]},"]
		(@past_e[count] ||= []) << ["#{row[9]},"]
		(@past_detail[count] ||= []) << ["#{row[11]},"]
		(@past_emotion[count] ||= []) << ["#{row[12]},"]
		(@past_valence[count] ||= []) << ["#{row[13]},"]
		(@past_significance[count] ||= []) << ["#{row[14]},"]
		file_past_c = File.new("xls_to_csv/past_c_#{count}.csv", "w+")
	 	print "PAST_C #{@past_c[count]}\n" 
         		for e in @past_c[count] do
  	    			file_past_c.write e
         		end #end for
         	file_past_c.close
		file_past_e = File.new("xls_to_csv/past_e_#{count}.csv", "w+")
	 	print "PAST_E #{@past_e[count]}\n" 
         		for e in @past_e[count] do
  	    			file_past_e.write e
         		end #end for
         	file_past_e.close
		file_past_detail = File.new("xls_to_csv/past_detail_#{count}.csv", "w+")
	 	print "past_DETAIL #{@past_detail[count]}\n" 
         		for e in @past_detail[count] do
  	    			file_past_detail.write e
         		end #end for
         	file_past_detail.close
		file_past_emotion = File.new("xls_to_csv/past_emotion _#{count}.csv", "w+")
	 	print "past_emotion  #{@past_emotion [count]}\n" 
         		for e in @past_emotion [count] do
  	    			file_past_emotion .write e
         		end #end for
         	file_past_emotion .close
		file_past_significance = File.new("xls_to_csv/past_significance _#{count}.csv", "w+")
	 	print "past_significance  #{@past_significance [count]}\n" 
         		for e in @past_significance [count] do
  	    			file_past_significance .write e
         		end
         	file_past_significance .close
		file_past_valence = File.new("xls_to_csv/past_valence _#{count}.csv", "w+")
	 	print "past_valence  #{@past_valence [count]}\n" 
         		for e in @past_valence [count] do
  	    			file_past_valence .write e
         		end #end for
         	file_past_valence .close
        end #end if
	#separate if-condition for "rating" because all conditions are relevant
	if row[0] == count
	(@rating[count] ||= []) << ["#{row[10]},"]
	file_rating = File.new("xls_to_csv/rating_#{count}.csv", "w+")
	print "RATING #{@rating[count]}\n" 
         	for e in @rating[count] do
  	    		file_rating.write e.join("\t\t\t")
         	end #end for
        file_rating.close
	end #end if

    	count += 1
     end #end while
  end #end book
end #end Spreadsheet


