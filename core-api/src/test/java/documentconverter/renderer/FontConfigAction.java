package documentconverter.renderer;

public class FontConfigAction {
	private String name;
	private float size;

	public FontConfigAction(String name, float size) {
		this.name = name;
		this.size = size;
	}

	public String getName() {
		return name;
	}

	public float getSize() {
		return size;
	}
}
