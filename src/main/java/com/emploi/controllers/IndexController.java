package com.emploi.controllers;


import java.io.File;
import java.io.FileInputStream;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import java.util.Iterator;

import java.util.ListIterator;
import java.util.Map;

import java.util.TreeMap;
import java.util.concurrent.TimeUnit;








import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

import org.openqa.selenium.phantomjs.PhantomJSDriver;





import javax.mail.MessagingException;


import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;




@Controller
public class IndexController {
	
	
   @Autowired
   private SmtpMailSender smtpMailSender;
	
	
	
	private ArrayList<String> villes = new ArrayList<String>();
	private ArrayList<Document> pagesDesProfiles = new ArrayList<Document>();
	private ArrayList<Document> listProfilesHtml = new ArrayList<Document>();
	private ArrayList<Document> listOffres = new ArrayList<Document>();
	private Document docum;
	private Document docVille;
	private Document docOffre;
    private  Map<String , String> mapMetiers = new TreeMap<>();
    private  Map<String , String> mapExperiences = new TreeMap<>();
    private  Map<String , String> mapLangues = new TreeMap<>();
    private  Map<String , String> mapMobiliteGeorgraphique= new TreeMap<>();
    private  Map<String , String> mapContrat= new TreeMap<>();
    private  Map<String , String> mapAnapecExperiences= new TreeMap<>();
    private  Map<String , String> mapAnapecLangues = new TreeMap<>();
    
    
    private ArrayList<String> listMetiers = new ArrayList<String>();
    private ArrayList<String> listExperiences = new ArrayList<String>();
    private ArrayList<String> listMobiliteGeographique = new ArrayList<String>();
    private ArrayList<String> listLangues = new ArrayList<String>(); 
    private ArrayList<String> listLanguesBD = new ArrayList<String>(); 
    private ArrayList<String> listContrat = new ArrayList<String>(); 
    private ArrayList<String> listAnapecExperiences = new ArrayList<String>();
    private ArrayList<String> listAnapecLangues = new ArrayList<String>(); 
    
    
    private String email;
    private String motdepasse;
	
    @RequestMapping("/")
	public String index() throws MessagingException {
    	
    	
		return "indexAuthentification";
	}
	
	
    	
	
	
@RequestMapping(value="/authentifier", method=RequestMethod.POST)
public String authentifier(@ModelAttribute("email") String email, @ModelAttribute("password") String password) throws IOException {

	this.email=email;
	this.motdepasse=password;
	
return "index";

	}

	
	@RequestMapping(value="/afficher", method=RequestMethod.POST)
	
	public String afficher(@ModelAttribute("fichier") String fichier,@ModelAttribute("email") String email, @ModelAttribute("DemandeOffre") String DemandeOffre,@ModelAttribute("ville") String ville,@ModelAttribute("de") int de,@ModelAttribute("A") int a , @ModelAttribute("pic") String pathOfExcel) throws IOException {
		
		
		if(DemandeOffre.equalsIgnoreCase("demandeEmploi")){
			run(de, a,ville , fichier, email);
		}
		else{
			runOffre(de, a);
		}
	   //return DemandeOffre +" : " + ville+" : "+ de+" : "+ a;
		return "index";
	}
	
	
	
	
public void run(int min , int ila, String ville, String fichier, String lemail) throws IOException {
	File f=new File(fichier+".xlsx");
	
		try {
			
			inialiserMetier();
	    	initialiserExperience();
	    	initialiserMapContrat();
	    	initialiserMapLangues();
	    	initialiserMapMobiliteGeo();
	    	listLanguesBD=listLangues;
	    	
		 String villeActuelle=ville;
		
		 
		   // File src=new File(pathOfExcel);
		   // FileInputStream fis = new FileInputStream(src);
			XSSFWorkbook wb = new XSSFWorkbook(); 
			XSSFSheet sheet1 = wb.createSheet("Chercheurs d'Emploi");
			//XSSFWorkbook wb = new XSSFWorkbook(fis);
			//SXSSFWorkbook wb = new SXSSFWorkbook();
			//XSSFSheet sheet1 = wb.getSheetAt(0);
			
			int indiceLigne = sheet1.getLastRowNum();

		   File file1=new File("phantomjs.exe");
	   	   System.setProperty("phantomjs.binary.path",file1.getAbsolutePath());

	   	   WebDriver driver=new PhantomJSDriver();
		    	String baseUrl = "https://www.emploi.ma/login";
		    	driver.get(baseUrl);
		    	driver.findElement(By.id("edit-name")).clear();
		    	driver.findElement(By.id("edit-name")).sendKeys(email);
		    	driver.findElement(By.id("edit-pass")).clear();
		    	driver.findElement(By.id("edit-pass")).sendKeys(motdepasse);
		    	driver.findElement(By.id("edit-submit")).click();
		    	driver.manage().timeouts().implicitlyWait(600, TimeUnit.SECONDS);
	            driver.manage().timeouts().pageLoadTimeout(600, TimeUnit.SECONDS);
	            
		 
		 String urll = "http://www.emploi.ma/recherche-base-donnees-cv/" + villeActuelle;
		 docVille= Jsoup.connect(urll+"?page=0").userAgent("Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1667.0 Safari/537.36").timeout(0).get();
		 Long nombrePage = Long.parseLong(docVille.getElementsByAttributeValue("class", "pager-current first").get(0).children().get(1).text());
		 int iP;
		 for(iP = min ; iP<=ila; iP++ ){ 
		  docVille= Jsoup.connect(urll+"?page="+iP).userAgent("Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1667.0 Safari/537.36").timeout(0).get();
          pagesDesProfiles.add(docVille);
          System.out.println(urll+"?page="+iP);    
		 
		   Iterator<Document> listIterator = pagesDesProfiles.iterator();
	       while(listIterator.hasNext()){	    	   
	        Document docVille = listIterator.next();
	    	Elements elts1 = docVille.getElementsByAttributeValue("class", "title search-title");
	    	
	    	for(Element elementProfile : elts1){
	
		    String numeroProfile = extractString(elementProfile.text(), 9);
		    Long numeroProfileLong = Long.parseLong(numeroProfile);
		    String urlProfile = "http://www.emploi.ma/recrutement-maroc-cv/"+ numeroProfileLong;
		    //docum = Jsoup.connect(urlProfile).userAgent("Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1667.0 Safari/537.36").timeout(0).get();
		   //driver.manage().timeouts().implicitlyWait(-1, TimeUnit.MINUTES);
		    driver.get(urlProfile);  
		    String htmlContent = driver.getPageSource();
	        docum = Jsoup.parse(htmlContent);
		    System.out.println(urlProfile);
		                                               
	       //driverMethode(urlProfilesLiens);
	       /*
	    	Iterator<String> it= urlProfilesLiens.iterator();
	    	while(it.hasNext()){
	    	String lienProfile = it.next();	
	    	driver.get(lienProfile);  
	        String htmlContent = driver.getPageSource();
	        docum = Jsoup.parse(htmlContent);
		    listProfilesHtml.add(docum);
	    	}
	    	*/
	       
	       // partie Net est terminé
	        
	        /*Iterator<Document> listIteratorProfiles = listProfilesHtml.iterator();
	        System.out.println("le nombre de profiles est : "+ listProfilesHtml.size());
	        System.out.println("la dernier ligne dans excel est : "+ indiceLigne);  */
	       
	       // while(listIteratorProfiles.hasNext()){
		    indiceLigne = sheet1.getLastRowNum();
	        	//System.out.println("la dernier ligne dans excel est : "+ indiceLigne);
		    indiceLigne++;
        	Row row = sheet1.createRow(indiceLigne);
	          //  Document docum=listIteratorProfiles.next();	
	        	String 	numeroProfile1= docum.getElementsByAttributeValue("class","group-title-inner inner").text();
	        	
	        	row.createCell(0).setCellValue(numeroProfile1);
	        	// System.out.println(numeroProfile1);
	        	
	        	row.createCell(1).setCellValue(villeActuelle);
	        
			     
			   //niveau d'etude bac+               
			    String a = docum.getElementsByAttributeValue("class", "candidate-education-info").get(0).getElementsByAttributeValue("class", "field-item").text();
			    String bac = extractString(a, 18);
			    row.createCell(2).setCellValue(bac);
			    
			    
			    String dateDebut=null, dateFin=null, title=null, etabl=null;
			    //Formations et periodes des formations
			    if(docum.getElementsByAttributeValue("class", "field-item field-candidate-formation").size()>=1){
			    Elements element = docum.getElementsByAttributeValue("class", "field-item field-candidate-formation");
			    
			    if(element.get(0).getElementsByAttributeValue("class", "emaxp-item default-format").size()>=1){
			    Elements elements = element.get(0).getElementsByAttributeValue("class", "emaxp-item default-format");
			    row.createCell(4).setCellValue("oui");
			    Element el = elements.first(); //prendre la derniere Formation
			    	String periode;
			    	//periode
			    	if(el.getElementsByAttributeValue("class", "period").size()>=1){
			    		periode = el.getElementsByAttributeValue("class", "period").text();
			    		String [] periodeDivise = periode.split(" - ");
			    		
			    		if(periodeDivise.length>=2){
			    			dateDebut = periodeDivise[0].replace(".", "/");
			    			dateFin = periodeDivise[1].replace(".", "/");
			    			row.createCell(5).setCellValue(dateDebut);
			    			row.createCell(6).setCellValue(dateFin); 
			    			
			    		}
			    		else if (periodeDivise.length== 1){
			    			row.createCell(5).setCellValue(periodeDivise[0]);
			    		}
			    		else {
			    			row.createCell(5).setCellValue(periode);
			    		     }
			    	}
				   //titre de la Formation
				   if(el.getElementsByAttributeValue("class", "title").size()>=1){
				   String titre = el.getElementsByAttributeValue("class", "title").text();
				   title=titre;
				   row.createCell(7).setCellValue(titre);
				    }
				   //etablissement
				   if(el.getElementsByAttributeValue("class", "establishment").size()>=1){
				   String etablissement = el.getElementsByAttributeValue("class", "establishment").text();
				  etabl=etablissement;
				  row.createCell(8).setCellValue(etablissement);
				    }
			    }
			    }
			    else{
			    	row.createCell(4).setCellValue("non");
			    }
			    
			    
			    
				  //Competences
				 String competences;
				 if(docum.getElementsByAttributeValue("class", "candidate-skills").get(0).text().length()<=1000){
			       competences = docum.getElementsByAttributeValue("class", "candidate-skills").get(0).text();
			       row.createCell(3).setCellValue(competences);
				  }
				 else{
					  competences = "non renseigné";
					  row.createCell(3).setCellValue(competences);
				     }
				 
				 Element elementsMoreInformations = docum.getElementsByAttributeValue("class", "candidate-more-info").get(0); 
				 Elements listMoreInformations = elementsMoreInformations.getElementsByAttributeValue("class", "field-items").get(0).getElementsByAttributeValue("class", "field-item");
				 String infor1= extractString(listMoreInformations.get(0).text(),16);
				 row.createCell(104).setCellValue(infor1); 
				 
				//experiences professionnelles et periodes
				 Elements elementsExepriences = docum.getElementsByAttributeValue("class", "candidate-professional-experience");
				 String niveauExperienc = elementsExepriences.get(0).getElementsByAttributeValue("class", "field-item").get(0).text();
				 niveauExperienc= extractString(niveauExperienc,22);
				 row.createCell(10).setCellValue(niveauExperienc);
				 

				 
				 //Langues
				 if(docum.getElementsByAttributeValue("class", "candidate-languages").size()>=1){
			      Element elementsLangues = docum.getElementsByAttributeValue("class", "candidate-languages").get(0);  
			      if(elementsLangues.getElementsByAttributeValue("class", "field-items").get(0).getElementsByAttributeValue("class", "language-item").size()>=1){
			      Elements listLangues = elementsLangues.getElementsByAttributeValue("class", "field-items").get(0).getElementsByAttributeValue("class", "language-item");
			      
			      for(Element el : listLangues){
				  String  Langue = extractString(el.getElementsByAttributeValue("class", "language-item").text(),2);
				  String [] listeLangue  = Langue.split(" / ");
				  
					if(listeLangue.length>=2){
						mapLangues.put(listeLangue[0], listeLangue[1]);
						
						//int idLangue = lgDao.findAll().indexOf((String)listeLangue[0]) + 6 ;
						
					    int idLangue =listLanguesBD.indexOf(listeLangue[0])+1;
						
						
					}
					} 
			      
			      affecterLanguesExcel(row, 97); 
			     
			      } }
				 
				 
				 //Plus d'informations
				   String infor2= extractString(listMoreInformations.get(1).attr("class", "regions-items").text(),24);
				   String infor3= extractString(listMoreInformations.get(2).attr("class", "job-type-items").text(),29);
				   
				   String [] infoGeographique = infor2.split(" - ");
				   String [] infoContrat = infor3.split(" - ");
				 
				   row.createCell(77).setCellValue("oui");
				 
				  
				   
						 for(String s : infoGeographique){
						 mapMobiliteGeorgraphique.put(s, "oui");
						// int  idVille = mbltgDao.findAll().indexOf(s) + 1;
						int  idVille = listMobiliteGeographique.indexOf(s)+1;
					
					   }
				  
						  affecterMobiliteGeoExcel(row, 78);
				   
				   
				   //informations Contrats
				   row.createCell(89).setCellValue("oui");
				   for(String s : infoContrat){
			          mapContrat.put(s, "oui");
			          int idContrat =listContrat.indexOf(s)+1;
						
			       
				   }
				   affecterContratExcel(row, 90); 
				   

			        	//Types de métiers recherchés
					     if(docum.getElementsByAttributeValue("class", "candidate-job-categories").get(0).getElementsByAttributeValue("class", "job-item").size()>=1){
						    Elements elementsMetiers = docum.getElementsByAttributeValue("class", "candidate-job-categories").get(0).getElementsByAttributeValue("class", "job-item");
						    row.createCell(62).setCellValue("oui");
						    for(Element el : elementsMetiers){
					        String metier = extractString(el.text(),2);
					        mapMetiers.put(metier, "oui"); 
					       // int  idMetier = mtrDao.findAll().indexOf(metier) + 1;
					        int  idMetier=  listMetiers.indexOf((String)metier)+1;
					      
					      }
						    affecterMetierExcel(row, 63);
						
					     } 
					     else{
					    	 row.createCell(62).setCellValue("non");  
					     }
				  
					     
					     
					//experiences professionnelles et periodes
					if(elementsExepriences.get(0).getElementsByAttributeValue("class", "field-item").size()>=2){
						Elements listSecteursExperience = elementsExepriences.get(0).getElementsByAttributeValue("class", "field-item").get(1).getElementsByAttributeValue("class", "industry-item");
						row.createCell(11).setCellValue("oui");
						
					    for(Element el : listSecteursExperience){
					    	
						String Secteur = extractString(el.getElementsByAttributeValue("class", "industry-item").text(),2);
					    mapExperiences.put(Secteur, "oui");
					   
					 // int  idSecteur = sctrDao.findAll().indexOf(Secteur) + 1;
					  int  idSecteur = listExperiences.indexOf(Secteur)+1;
					   
						}
					    
					    affecterExperiencesExcel(row, 12);
					}
					else{
						row.createCell(11).setCellValue("non");
						affecterExperiencesExcel(row, 12);
					}
				
				  
			        mapContrat.clear();
			        mapExperiences.clear();
			        mapLangues.clear();
			        mapMetiers.clear();
			        mapMobiliteGeorgraphique.clear();
			        
			     //System.out.println("le numero de profile : "+ numeroProfile + "  est terminé");
			        
			        FileOutputStream out = new FileOutputStream(f);
			        
		            wb.write(out);
		            out.close();
		            
		            /*
			        FileOutputStream fout = new FileOutputStream(src);
			        wb.write(fout);*/
			       
	        }
	    	elts1.clear();
	    	}
	       
	      
	     
	        //vider les listes
	        pagesDesProfiles.clear();
	        listProfilesHtml.clear();
	        }//fermeture de for
		 
		driver.close();
    	driver.quit();
		  
		
	      
    	wb.close();
	       
	        listMetiers.clear(); 
	        listExperiences.clear();
	        listMobiliteGeographique.clear();
	        listLangues.clear();
	        listContrat.clear();  
	        
	    smtpMailSender.send(lemail, "Test mail", "mazyann", f);    
	       
		} catch (Exception e) {
			
			e.printStackTrace();
		}
		
		
	}// férmeture de la methode run 

	
	





















public void runOffre(int min , int ila) throws IOException {
	
	try {	
	//initialiser tous les listes
    inialiserMetier();
	initialiserExperience();
	initialiserMapContrat();
	initialiserMapLangues();
	initialiserMapMobiliteGeo();
	
	
	File src=new File("C:\\Users\\BEN AHMID Soufiane\\Documents\\PFE de 2016\\DB_Offre_Emploi.xlsx");
	FileInputStream fis = new FileInputStream(src);
	XSSFWorkbook wb = new XSSFWorkbook(fis);
	XSSFSheet sheet1 = wb.getSheetAt(0);
	int indiceLigne = sheet1.getLastRowNum();
	/*
	int jr=sheet1.getRow(indiceLigne).getCell(2).getDateCellValue().getDay();
	int ms=sheet1.getRow(indiceLigne).getCell(2).getDateCellValue().getMonth();
	int ane=sheet1.getRow(indiceLigne).getCell(2).getDateCellValue().getYear();
	Date da1 = new Date(jr, ms, ane);
	
	jr=sheet1.getRow(1).getCell(2).getDateCellValue().getDay();
    ms=sheet1.getRow(1).getCell(2).getDateCellValue().getMonth();
    ane=sheet1.getRow(1).getCell(2).getDateCellValue().getYear();
	Date da2 = new Date(jr, ms, ane);
	
	Date date1;
	
	if(da1.compareTo(da2)==1){
		date1=da1;
		}
	else{
		date1=da2;
	}
	*/
	System.out.println("la dernier ligne est : "+ indiceLigne);
	
	for(int k=min; k<=ila; k++){ // on extrait page par page
		
	
	docOffre = Jsoup.connect("http://www.emploi.ma/recherche-jobs-maroc?page="+k).userAgent("Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1667.0 Safari/537.36").timeout(0).get();
	
	if(docOffre.getElementsByAttributeValue("class", "pager-current first").size()>=1){
	Long nombrePageOffre = Long.parseLong(docOffre.getElementsByAttributeValue("class", "pager-current first").get(0).children().get(1).text());
	}
	
	Elements elementsOffre = docOffre.getElementsByAttributeValue("class", "title search-title");
	for(Element el : elementsOffre){
	String urlOffre = el.select("a").first().attr("href");	
	System.out.println(urlOffre);
	Document d =Jsoup.connect("http://www.emploi.ma"+urlOffre).userAgent("Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1667.0 Safari/537.36").timeout(0).get();
	listOffres.add(d);
	}
	
	ListIterator<Document> listIteratorOffres = listOffres.listIterator();
	
	while(listIteratorOffres.hasNext()){// on parcour la liste des profiles 
		
		indiceLigne++;
    	Row rowOffre = sheet1.createRow(indiceLigne);
		Document d = listIteratorOffres.next();
		
		String datePublication=null;
		if(d.getElementsByAttributeValue("class", "job-ad-publication-date").size()>=1){
			datePublication =  extractString(d.getElementsByAttributeValue("class", "job-ad-publication-date").get(0).text(),13).replace(".", "/"); 
			String[] tabDatePublication = datePublication.split("/");
			int jour = Integer.parseInt(tabDatePublication[0]);
			int mois = Integer.parseInt(tabDatePublication[1]);
			int annee = Integer.parseInt(tabDatePublication[2]);	
			
			/*
			 if(date2.compareTo(date1)==-1 ||  date2.compareTo(date1)==0){
				System.exit(0);
			}
			*/
			
			rowOffre.createCell(2).setCellValue(datePublication);
			
			}
		
		String nomEntreprise=null;
		if(d.getElementsByAttributeValue("class", "company-title").size()>=1){
		nomEntreprise = d.getElementsByAttributeValue("class", "company-title").get(0).text();
		rowOffre.createCell(0).setCellValue(nomEntreprise);
		//System.out.println("les nom  de l'entreprise est : "+ nomEntreprise);
		}
		
		String postePropose=null;
		if(d.getElementsByAttributeValue("id", "group-title-inner").size()>=1){
		postePropose = d.getElementsByAttributeValue("id", "group-title-inner").get(0).text();
		rowOffre.createCell(1).setCellValue(postePropose);
		}
		
		
		
		String niveauExperience=null;
		String niveauEtude=null;
		String contrat=null;
		ArrayList<String> tousLesSecteurs = new ArrayList<String>();
		ArrayList<String> tousLesMetiers = new ArrayList<String>();
		ArrayList<String> tousLesRegions = new ArrayList<String>();
		ArrayList<String> tousLesLangues = new ArrayList<String>();
		ArrayList<String> tousLesContrats = new ArrayList<String>();
		
		Elements elementsTr = d.getElementsByAttributeValue("class", "job-ad-criteria").get(0).select("tr");
		for(Element el : elementsTr ){
			
		String titre = el.select("td").first().text();
		
		if(titre.equalsIgnoreCase("Métier") ){
			rowOffre.createCell(32).setCellValue("oui");
		     String [] tableau = el.select("td").last().html().toString().split("<br>");
		     int i=0;
		     for(String s : tableau){
		    if(i==0){
		     mapMetiers.put(s, "oui");
		     tousLesMetiers.add(s);
		    	 }
		    
		    else{
		    	String mtr=extractString(s, 1);
		    	mapMetiers.put(mtr, "oui"); 
		    	tousLesMetiers.add(mtr);
		    }
		    i++;
		}
		
		}
		
		else if(titre.equalsIgnoreCase("Secteur d´activité")){
			rowOffre.createCell(47).setCellValue("oui");
			 String [] tableau = el.select("td").last().html().split("<br>");
			 int i=0;
			 for(String s : tableau){
			 if(i==0){
			 mapExperiences.put(s, "oui");
			 tousLesSecteurs.add(s);
				}
			    else{
			    	String sct=extractString(s, 1);
			    	mapExperiences.put(sct, "oui");
			    	tousLesSecteurs.add(sct);
			    }
			    i++;
		}
		}
		
		else if(titre.equalsIgnoreCase("Type de contrat")){
			 rowOffre.createCell(15).setCellValue("oui");
			 contrat = el.select("td").last().text();
			 mapContrat.put(contrat, "oui");
			 
		}
				else if(titre.equalsIgnoreCase("Région")){
			rowOffre.createCell(3).setCellValue("oui");
		String [] tableau =	 el.select("td").last().text().split(",");
		 int i=0;
		for(String s : tableau){
			if(i==0){
			mapMobiliteGeorgraphique.put(s, "oui");
			tousLesRegions.add(s);
			}
		    else{
		    	String rgn = extractString(s, 1);
		    	mapMobiliteGeorgraphique.put(rgn, "oui"); 
		    	tousLesRegions.add(rgn);
		    }
		    i++;
		}
		}
		
		else if(titre.equalsIgnoreCase("Niveau d'expérience")){
			niveauExperience = el.select("td").last().text();
			rowOffre.createCell(31).setCellValue(niveauExperience);
		}
		else if(titre.equalsIgnoreCase("Niveau d'études")){
			niveauEtude = el.select("td").last().text();
			rowOffre.createCell(30).setCellValue(niveauEtude);
		}
		else if(titre.equalsIgnoreCase("Langues exigées")){
			
			String  phrase = el.select("td").last().text().replace("(", "- ").replace(")", " -");
			//System.out.println(phrase);
			String [] tableau = phrase.split(" - ");
			
			if(tableau.length>=2){
			for(int i=0; i<tableau.length; ){
				String nivLang = tableau[i+1].replace("-", "");
				mapLangues.put(tableau[i], nivLang);
				tousLesLangues.add(tableau[i]);
				tousLesLangues.add(nivLang);
				i=i+2;
			}
			}
		}
		
		}
		
		
		
		
		Iterator<String> itererTousLesSecteurs = tousLesSecteurs.listIterator();
		while(itererTousLesSecteurs.hasNext()){
			String leSecteur=itererTousLesSecteurs.next();
			int  idSecteur = listExperiences.indexOf(leSecteur)+1;
		  
		}
		
		Iterator<String> itererTousLesMetiers = tousLesMetiers.listIterator();
		while(itererTousLesMetiers.hasNext()){
	    String leMetier=itererTousLesMetiers.next();
		int  idMetier=  listMetiers.indexOf((String)leMetier)+1;
    
		}
		 		
		
		Iterator<String> itererTousLesRegions = tousLesRegions.listIterator();
		while(itererTousLesRegions.hasNext()){
	    String laRegion=itererTousLesRegions.next();
	    int  idVille = listMobiliteGeographique.indexOf(laRegion)+1;
		
		}
		
		
		Iterator<String> itererTousLesLangues = tousLesLangues.listIterator();
		while(itererTousLesLangues.hasNext()){
	    String laLangue=itererTousLesLangues.next();
	    String leNiveau=itererTousLesLangues.next();
	    int idLangue =listLangues.indexOf(laLangue)+1;
	
		}
		
		
		int idContrat =listContrat.indexOf(contrat)+1;
      
		
		affecterContratExcel(rowOffre, 16);
		affecterExperiencesExcel(rowOffre, 48);
		affecterLanguesOffreExcel(rowOffre, 23);
		affecterMetierExcel(rowOffre, 33);
		affecterMobiliteGeoExcel(rowOffre, 4);
		
		
		    FileOutputStream fout = new FileOutputStream(src);
	        wb.write(fout);	
	        mapContrat.clear();
	        mapExperiences.clear();
	        mapLangues.clear();
	        mapMetiers.clear();
	        mapMobiliteGeorgraphique.clear();
	        tousLesLangues.clear();
	        tousLesRegions.clear();
	        tousLesMetiers.clear();
	        tousLesSecteurs.clear();
	        
	}
	
	listOffres.clear();
	}     
        wb.close();
       
        listMetiers.clear(); 
        listExperiences.clear();
        listMobiliteGeographique.clear();
        listLangues.clear();
        listContrat.clear();  
       
	} catch (Exception e) {
		
		e.printStackTrace();
	}
	
	
}// férmeture de la methode runOffre 





































public void inialiserMetier(){
    
    listMetiers.add("Achats, transport, logistique");
	listMetiers.add("Commercial, vente");
    listMetiers.add("Gestion, comptabilité, finance");
    listMetiers.add("Informatique, nouvelles technologies");
    listMetiers.add("Management, direction générale");
    listMetiers.add("Marketing, communication");
    listMetiers.add("Métiers de la santé et du social");
    listMetiers.add("Métiers des services"); 
    listMetiers.add("Métiers du BTP"); 
    listMetiers.add("Production, maintenance, qualité"); 
    listMetiers.add("R&D, gestion de projets"); 
    listMetiers.add("RH, juridique, formation"); 
    listMetiers.add("Secrétariat, assistanat");
    listMetiers.add("Tourisme, hôtellerie, restauration");
  
}

public void initialiserExperience(){


listExperiences.add("Activités associatives");
listExperiences.add("Administration publique");
listExperiences.add("Aéronautique, navale");
listExperiences.add("Agriculture, pêche, aquaculture");
listExperiences.add("Agroalimentaire");
listExperiences.add("Ameublement, décoration");
listExperiences.add("Automobile, matériels de transport, réparation");
listExperiences.add("Banque, assurance, finances");
listExperiences.add("BTP, construction");
listExperiences.add("Centres d´appels, hotline, call center");
listExperiences.add("Chimie, pétrochimie, matières premières");
listExperiences.add("Conseil, audit, comptabilité");
listExperiences.add("Distribution, vente, commerce de gros");
listExperiences.add("Édition, imprimerie");
listExperiences.add("Éducation, formation");
listExperiences.add("Electricité, eau, gaz, nucléaire, énergie");
listExperiences.add("Environnement, recyclage");
listExperiences.add("Equip. électriques, électroniques, optiques, précision");
listExperiences.add("Equipements mécaniques, machines");
listExperiences.add("Espaces verts, forêts, chasse");
listExperiences.add("Événementiel, hôte(sse), accueil");
listExperiences.add("Hôtellerie, restauration");
listExperiences.add("Immobilier, architecture, urbanisme");
listExperiences.add("Import, export");
listExperiences.add("Industrie pharmaceutique");
listExperiences.add("Industrie, production, fabrication, autres");
listExperiences.add("Informatique, SSII, Internet");
listExperiences.add("Ingénierie, études développement");
listExperiences.add("Intérim, recrutement");
listExperiences.add("Luxe, cosmétiques");
listExperiences.add("Emplacement");
listExperiences.add("Maintenance, entretien, service après vente");
listExperiences.add("Manutention");
listExperiences.add("Marketing, communication, médias");
listExperiences.add("Métallurgie, sidérurgie");
listExperiences.add("Nettoyage, sécurité, surveillance");
listExperiences.add("Papier, bois, caoutchouc, plastique, verre, tabac");
listExperiences.add("Produits de grande consommation");
listExperiences.add("Qualité, méthodes");
listExperiences.add("Recherche et développement");
listExperiences.add("Santé, pharmacie, hôpitaux, équipements médicaux");
listExperiences.add("Secrétariat");
listExperiences.add("Services autres");
listExperiences.add("Services aéroportuaires et maritimes");
listExperiences.add("Services collectifs et sociaux, services à la personne");
listExperiences.add("Sport, action culturelle et sociale");
listExperiences.add("Textile, habillement, cuir, chaussures");
listExperiences.add("Tourisme, loisirs");
listExperiences.add("Transports, logistique, services postaux");
listExperiences.add("Télécom");

}

public void initialiserMapMobiliteGeo(){

listMobiliteGeographique.add("Agadir");
listMobiliteGeographique.add("Casablanca");
listMobiliteGeographique.add("Fès");
listMobiliteGeographique.add("Laâyoune");
listMobiliteGeographique.add("Marrakech");
listMobiliteGeographique.add("Meknès");
listMobiliteGeographique.add("Oujda");
listMobiliteGeographique.add("Rabat");
listMobiliteGeographique.add("Settat");
listMobiliteGeographique.add("Tanger");
listMobiliteGeographique.add("International");
}

public void initialiserMapLangues(){
listLangues.add("arabe");
listLangues.add("français");
listLangues.add("anglais");
listLangues.add("espagnol");
listLangues.add("allemand");
listLangues.add("italien");
listLangues.add("berbère");
}

public void initialiserMapContrat(){

listContrat.add("CDI");
listContrat.add("CDD");
listContrat.add("Intérim");
listContrat.add("Stage");
listContrat.add("Freelance");
listContrat.add("Alternance");
listContrat.add("Temps partiel");

}


public void affecterContratExcel(Row r, int i){
if(!mapContrat.isEmpty()){
int index=i;  //90
//Set<String> ensemble = mapContrat.keySet();
Iterator<String> it = listContrat.iterator();

while(it.hasNext()){
	String s = it.next();
	
	if(mapContrat.get(s)!=null){
	r.createCell(index).setCellValue(mapContrat.get(s));
	}
	else{
	r.createCell(index).setCellValue("non");	
	}
	
	 index++;
	 
}

}
else{
r.createCell(i-1).setCellValue("non");
}

}


public void affecterMobiliteGeoExcel(Row r, int i){

if(!mapMobiliteGeorgraphique.isEmpty()){
int index=i; //78
//Set<String> ensemble = mapMobiliteGeorgraphique.keySet();
Iterator<String> it = listMobiliteGeographique.iterator();

while(it.hasNext()){
	String s = it.next();
	if(mapMobiliteGeorgraphique.get(s)!=null){
	 r.createCell(index).setCellValue(mapMobiliteGeorgraphique.get(s));
	}
	else{
	 r.createCell(index).setCellValue("non");	
	}
	 index++; 
}
}
else{
	r.createCell(i-1).setCellValue("non");
}


}



public void affecterLanguesExcel(Row r,  int i ){

int index= i; //97
//Set<String> ensemble = mapLangues.keySet();
Iterator<String> it = listLangues.iterator();

while(it.hasNext()){
	
	String s = it.next();
	if(mapLangues.get(s)!=null){
		r.createCell(index).setCellValue(mapLangues.get(s));
	}
	else{
	 r.createCell(index).setCellValue("non parlé");	
	}
	 index++;
	 
}
	 
}

public void affecterLanguesOffreExcel(Row r,  int i ){

int index= i; //97
//Set<String> ensemble = mapLangues.keySet();
Iterator<String> it = listLangues.iterator();

while(it.hasNext()){
	
	String s = it.next();
	if(mapLangues.get(s)!=null){
		r.createCell(index).setCellValue(mapLangues.get(s));
	}
	else{
	 r.createCell(index).setCellValue("non exigé");	
	}
	 index++;
	 
}
	 
}

public void affecterMetierExcel(Row r, int i){

if(!mapMetiers.isEmpty()){
int index= i; // 63
//Set<String> ensemble = mapMetiers.keySet();
Iterator<String> it = listMetiers.iterator();

while(it.hasNext()){
	
	String s = it.next();
	if(mapMetiers.get(s)!=null){
	r.createCell(index).setCellValue(mapMetiers.get(s));
	}
	else{
	r.createCell(index).setCellValue("non");	
	}
	 index++;
}
}
else{
	r.createCell(i-1).setCellValue("non");
	Iterator<String> it = listMetiers.iterator();
	while(it.hasNext()){
	it.next();
	r.createCell(i).setCellValue("non");
	i++;
	}
}
}



public void affecterExperiencesExcel(Row r, int i){

if(!mapExperiences.isEmpty()){
int index= i;  //12
//Set<String> ensemble = mapExperiences.keySet();
Iterator<String> it = listExperiences.iterator();

while(it.hasNext()){
	
	String s = it.next();
	if(mapExperiences.get(s)!=null){
		r.createCell(index).setCellValue(mapExperiences.get(s));
	}
	else{
	r.createCell(index).setCellValue("non");	
	}

	 index++;
	 
	 
}
}
else{
	r.createCell(i-1).setCellValue("non");
	
	Iterator<String> it = listExperiences.iterator();
	while(it.hasNext()){
	it.next();
	r.createCell(i).setCellValue("non");
	i++;
	}
	
}	
}


private  String extractString(String text, int n){
if(text != null){

	String niveau = text.substring(n);
	
	return niveau;
}
return null;
}


	 
}
