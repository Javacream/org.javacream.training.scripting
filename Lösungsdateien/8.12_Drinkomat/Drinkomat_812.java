import java.util.Scanner;

/* HIER BEGINNT DIE KLASSENDEFINITION */
public class Drinkomat_812 {
    
    /* MAIN-METHODE WIRD BEIM STARTEN AUSGEF‹HRT */
    public static void main(String[] args) {
        
        /* VARIABLENDEKLARATION */
        int input;
        String drink;
        double charge;
        
        /* VARIABLENINITIALISIERUNG */
        input = 0;
        charge = 0;
        drink = "";
        
        /* WIRD BEN÷TIGT UM BEFEHLSZEILE EINZULESEN */
        Scanner scanner = new Scanner(System.in);
 
        /* MELDUNG AUSGEBEN */
        System.out.println("Willkommen beim Drinkomat! Bitte Getr√§nkenummer w√§hlen");
        
        /* EINGABE */  
        input = scanner.nextInt();
        System.out.println("Sie w√ºnschen: Getr√§nk " + input);
        
        /* PREIS UND NAME FESTLEGEN */
        switch (input) {
          case 1: 
            charge = 50;
            drink = "Apfelsaft";
            break;
          case 2: 
            charge = 75;
            drink = "Orangensaft";
            break;
          case 3:
            charge = 30;
            drink = "Mineralwasser";
            break;
          case 4:
            charge = 70;
            drink = "Cappucino";
            break;
          case 5:
            charge = 60;
            drink = "Espresso";
            break;
          default: 
            System.out.println("Falsche Getr√§nkenummer!");
            System.exit(0);
        }
        
        /* GELDEINWURF */
        System.out.println("Bitte werfen Sie: " + charge + " Cent ein.");
        input = scanner.nextInt();
        System.out.println("eingeworfen: " + input);
        
        /* GELDEINGABE PR‹FEN */
        if (input == charge)
        {
          /*GETR√NK AUSGEBEN */  
          System.out.println("Hier kommt Ihr " + drink + "!");
        }
        else
        {
          System.out.println("Falscher Betrag!");
        }

        scanner.close();
    }  
}
