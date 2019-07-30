import java.util.Scanner;

/* Hier beginnt die Klassendefinition */
public class Drinkomat_810 {
    
    /* Die main-Methode markiert den Programmstart */
    public static void main(String[] args) {
        
        /* Variablendeklaration */
        int input;
        String drink;
        int charge;
        
        /* Variableninitialisierung*/
        input = 0;
        charge = 0;
        drink = "";
        
        /* Vorbereitung f�r Eingaben */
       Scanner scanner = new Scanner(System.in);
 
        /* Meldung ausgeben */
        System.out.println("Willkommen beim Drinkomat! Bitte Getr�nkenummer w�hlen");
        
        /* Wingabe */  
        input = scanner.nextInt();
        System.out.println("Sie w�nschen: Getr�nk Nr. " + input);
        
        /* Preis und Namen festlegen */
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
          default: 
            charge = 50;
            drink = "Apfelsaft";
        }
        
        /* Geldeinwurf */
        System.out.println("Bitte werfen Sie " + charge + " Cent ein.");
        input = scanner.nextInt();
        System.out.println("eingeworfen: " + input);

        /* Geldeingabe pr�fen */
        if (input == charge) {
           /* Getr�nk ausgeben  */  
           System.out.println("Hier kommt Ihr " + drink + "!");
        } else {
           System.out.println("Falscher Betrag!");
        }

        scanner.close();
    }  
}
