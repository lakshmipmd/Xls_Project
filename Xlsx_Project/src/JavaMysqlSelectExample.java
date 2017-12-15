
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.io.File;
import java.io.FileNotFoundException;
import java.sql.Connection;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Scanner;

public class JavaMysqlSelectExample {

	private static final File txtFile = new File("C:\\Users\\pradp\\Documents\\pdp_fileproject_output\\query.txt");

	public static void main(String[] argv) {

		System.out.println("-------- Oracle JDBC Connection Testing ------");

		try {
			Class.forName("oracle.jdbc.driver.OracleDriver");
		} 
		catch (ClassNotFoundException e) {
			System.out.println("Where is your Oracle JDBC Driver?");
			e.printStackTrace();
			return;
		}

		System.out.println("Oracle JDBC Driver Registered!");

		Connection connection = null;
		try {
			connection = DriverManager.getConnection("jdbc:oracle:thin:@localhost:1521:orcl", "LPUSER", "C0ntr0ller");
		} catch (SQLException e) {
			System.out.println("Connection Failed! Check output console");
			e.printStackTrace();
			return;
		}

		if (connection != null) {
			System.out.println("You made it, take control your database now!");
		} else {
			System.out.println("Failed to make connection!");
		}

		System.out.println("Execute query from text file");
		try {
			Scanner input = new Scanner(txtFile);
			while (input.hasNext()) {
				String line = input.nextLine();
				System.out.println(line);
				Statement stmt = connection.createStatement();
				// execute insert SQL Statement
				stmt.executeUpdate(line);
			}
			System.out.println("Records are inserted into the table!");
			input.close();
		} catch (FileNotFoundException | SQLException e) {
			e.printStackTrace();
		}
		
		System.out.println("Execute the select statement");

		ResultSet rs;
		try {
			Statement stmt = connection.createStatement();

			rs = stmt.executeQuery("SELECT * FROM SYS.SAMPLE_TABLE");
			while (rs.next())
				System.out.println(rs.getInt(1) + "  " + rs.getString(2) + "  " + rs.getString(3) + "  "
						+ rs.getString(4) + "  " + rs.getString(5));

		} catch (SQLException e) {
			e.printStackTrace();
		}

	}

}