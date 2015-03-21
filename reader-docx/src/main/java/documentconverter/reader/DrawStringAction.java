package documentconverter.reader;

public class DrawStringAction {
	private String text;
	private int x;

	public DrawStringAction(String text, int x) {
		this.text = text;
		this.x = x;
	}

	public String getText() {
		return text;
	}
	
	public int getX() {
		return x;
	}
}
