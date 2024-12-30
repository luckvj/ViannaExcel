/*
Program Created by: Vincent Haney Jr.
Date: 02/25/2024
 */

import javax.swing.JOptionPane;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        try {
            String folderPath = System.getProperty("user.home") + "/Desktop/Real Estate/";
            java.io.File folder = new java.io.File(folderPath);
            if (!folder.exists()) {
                folder.mkdirs(); // Create the folder and any necessary parent folders
            }

            String messageLover = "This is a basic program that will develop overtime. To Savannah, Love Vincent.";
            String programVersion = "Version Pre-Alpha 0.01";
            JOptionPane.showMessageDialog(null, messageLover);
            JOptionPane.showMessageDialog(null, programVersion);              
            
            
            
            // Ask user if they have any property details to enter
            int option = JOptionPane.showConfirmDialog(null, "Do you have property details to enter?", "Input Confirmation", JOptionPane.YES_NO_OPTION);
            String fullName = JOptionPane.showInputDialog("Enter your client's full name:");          
            
            while (option == JOptionPane.YES_OPTION) {
            // Prompt user to enter property details
                String address = JOptionPane.showInputDialog("Enter your client's showing address:");
                String date = JOptionPane.showInputDialog("Enter the date you want to view the showing: ");
                String time = JOptionPane.showInputDialog("Enter the time you want to view the showing: ");
                String price = JOptionPane.showInputDialog("Enter the price of the building: ");
                int numBedrooms = Integer.parseInt(JOptionPane.showInputDialog("Enter Number of Bedrooms: "));
                int numBathrooms = Integer.parseInt(JOptionPane.showInputDialog("Enter Number of Bathrooms: "));

                // Specify the file name
                String fileName = fullName + "_PropertyDetails.xlsx";
                String filePath = folderPath + fileName;

                // Write property details to Excel file
                ExcelWriter.writeToExcel(fullName, address, date, time, price, numBedrooms, numBathrooms, folderPath + fileName);

                // Ask if the user has more inputs
                option = JOptionPane.showConfirmDialog(null, "Do you have more property details to enter?", "Input Confirmation", JOptionPane.YES_NO_OPTION);
            }

            // Show success message
            JOptionPane.showMessageDialog(null, "All property details have been entered successfully!");
            JOptionPane.showMessageDialog(null, "Excel file saved to Real Estate Folder on Desktop");
            String programCreator = "Program created by your's truly, Vincent.";
            JOptionPane.showMessageDialog(null, programCreator);
            
        } catch (IOException e) {
            // Show error message if an exception occurs
            JOptionPane.showMessageDialog(null, "Error occurred: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            e.printStackTrace();
        }
    }
}


