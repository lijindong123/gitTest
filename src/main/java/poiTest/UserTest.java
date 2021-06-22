package poiTest;

import java.math.BigDecimal;

public class UserTest {
	
	private int id;
	private String name;
	private int age;
	private String addr;
	private BigDecimal phone;
	public int getId() {
		return id;
	}
	public void setId(int id) {
		this.id = id;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public int getAge() {
		return age;
	}
	public void setAge(int age) {
		this.age = age;
	}
	public String getAddr() {
		return addr;
	}
	public void setAddr(String addr) {
		this.addr = addr;
	}
	public BigDecimal getPhone() {
		return phone;
	}
	public void setPhone(BigDecimal phone) {
		this.phone = phone;
	}
	@Override
	public String toString() {
		return "UserTest [id=" + id + ", name=" + name + ", age=" + age + ", addr=" + addr + ", phone=" + phone + "]";
	}
	public UserTest(int id, String name, int age, String addr, BigDecimal phone) {
		super();
		this.id = id;
		this.name = name;
		this.age = age;
		this.addr = addr;
		this.phone = phone;
	}
	
	
	
}
