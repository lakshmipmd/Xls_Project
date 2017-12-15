
public class ExcelFile {

	private double ID;
	private String Name;
	private String Salary;
	private String Department;
	private String Manager;
	
	public ExcelFile(){}
	
	public ExcelFile(double ID, String Name, String Salary, String Department, String Manager){
		this.ID = ID;
		this.Name = Name;
		this.Salary = Salary;
		this.Department = Department;
		this.Manager = Manager;
	}
	
	@Override
	public String toString(){
		return ID + "_" + Name + "_" + Salary + "_" + Department + "_" + Manager;
	}
	public double getID() {
		return ID;
	}
	public void setID(double iD) {
		ID = iD;
	}
	public String getName() {
		return Name;
	}
	public void setName(String name) {
		Name = name;
	}
	public String getSalary() {
		return Salary;
	}
	public void setSalary(String salary) {
		Salary = salary;
	}
	public String getDepartment() {
		return Department;
	}
	public void setDepartment(String department) {
		Department = department;
	}
	public String getManager() {
		return Manager;
	}
	public void setManager(String manager) {
		Manager = manager;
	}
	
}
