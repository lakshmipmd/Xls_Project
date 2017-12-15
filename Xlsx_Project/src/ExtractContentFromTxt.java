import java.io.File;
import java.io.FileNotFoundException;
import java.util.Scanner;

public class ExtractContentFromTxt {

	private static final File txtFile = new File("C:\\Users\\pradp\\Documents\\pdp_fileproject_output\\query.txt");

	public static void main(String[] args) {
		try {
			Scanner input = new Scanner(txtFile);
			while (input.hasNext()) {
				String line = input.nextLine();
				System.out.println(line);
			}
			input.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
