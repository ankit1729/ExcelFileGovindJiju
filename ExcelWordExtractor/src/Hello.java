
public class Hello {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		String input="You hello, hsdjj hello, ghdsjhdsjds ksakjdj hello jhfd";
		int ind=input.indexOf("hello");
		while(ind!=-1){
			System.out.println(ind);
			ind=input.indexOf("hello", ind+1);
		}

	}

}
