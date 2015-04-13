/*
 * Copyright (C) 2015 James Carter
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package ooxmltojava2d.docx;

public class DrawImageAction {
	private String relationshipId;
	private int width;
	private int height;
	private int x;

	public DrawImageAction(String relationshipId, int width, int height, int x) {
		this.relationshipId = relationshipId;
		this.width = width;
		this.height = height;
		this.x = x;
	}

	public String getRelationshipId() {
		return relationshipId;
	}

	public int getX() {
		return x;
	}

	public int getWidth() {
		return width;
	}

	public int getHeight() {
		return height;
	}
}
