require 'spreadsheet'  
require 'gmail' 
require "prawn" 

gmail = Gmail.connect("email@gmail.com", "password")
workbook = Spreadsheet.open 'E:\epson\ins.xls'
a = 0
sheet1 = workbook.worksheet 0
sheet1.each do |row|
        a = a + 1
        #puts "#{row[0]} , #{row[1]} , #{row[2]}"
        puts gmail.inbox.count
        puts gmail.inbox.count(:unread)
        puts gmail.inbox.count(:read)
        gmail.deliver do
             to "#{row[2]}"
             subject "Invitation de la formation CISCO"
             text_part do
                body "invitation CISCO. "
                #puts "nom:#{row[0]} et prenom:#{row[1]} et ID:#{a}"
                puts "yes send name #{row[2]}"
             end
             html_part do
             	
                 content_type 'text/html; charset=UTF-8'
                 body "<p>Cher(e) postulant(e),<br>

Nous avons le plaisir de vous informer que vous avez 
été pris pour prendre part à la formation CISCO « Implementing 
IP Switched & Secured Networks » qui se déroulera <h3>le 24, 25 et 
26 Décembre 2018 </h3>à l’auditorium de l’université des sciences 
et de la technologie Houari Bouemedienne à partir de <h3>8h00 
jusqu'à 16h00.</h3><br> 

Vous devrez être munis chaque jour de formation avec:<br>

Votre invitation correspondante a ce jour. 
( Les invitations sont en pièce jointe)<br> 

Votre pièce d'identité avec laquelle vous vous êtes inscrit.<br> 

Au courant de la formation, vous aurez certainement besoin
 d’un ordinateur portable,
 nous vous conseillons donc d’en ramener un avec vous.<br>

Nous comptons sur votre ponctualité.Veuillez vérifier vos noms
 et prénoms dans l'invitation, si toute fois il y'a une erreur 
 envoyez une réclamation a cet e-mail le plutôt possible.<br>   

Veuillez agréer nos salutations les plus distinguées.<br>

 

<h2>Attention!</h2><br>
 
<p>Sachez que l'invitation est unique. A ne pas dupliquer </p><br>

NB : La restauration et l'hébergement ne seront pas pris en charge.<br></p>"
                 puts "yes send information #{row[2]}"
             end

             ####################     pdf       ##################

             Prawn::Document.generate("E:\\disk-script\\rb-py\\CelecInvitationCisco\\Club.pdf") do     
 
                  pigs = "#{Prawn::DATADIR}/d123-2.jpg"
                  image pigs , :width => 500, :height => 700  
         
                  #################### day 1 #####
                  bounding_box([325, 583], :width => 300, :height => 400) do
                        text "#{row[0]} #{row[1]}"
                  end
                  bounding_box([453, 511], :width => 300, :height => 400) do
                        text "#{a}"
                  end
                  ####################  day 2 ######
                  bounding_box([325, 353], :width => 300, :height => 400) do
                        text "#{row[0]} #{row[1]}"
               
                  end
                  bounding_box([453, 280], :width => 300, :height => 400) do
                        text "#{a}"
                  end
                  ####################  day 3 ####
                  bounding_box([325, 115], :width => 300, :height => 400) do
                        text "#{row[0]} #{row[1]}"
                  end
                  bounding_box([453, 42], :width => 300, :height => 400) do
                        text "#{a}"
                  end
                  #################################
             end 

             ############################################
             #puts "s send file"
             puts "yes send pdf #{a} :1"
             add_file "E:\\disk-script\\rb-py\\CelecInvitationCisco\\Club.pdf"
             puts "yes send pdf #{a} :2"
             puts "*******************************"
  
        end


end
gmail.logout
puts "close" 