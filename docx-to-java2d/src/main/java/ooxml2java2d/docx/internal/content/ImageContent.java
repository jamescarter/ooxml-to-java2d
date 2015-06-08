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

package ooxml2java2d.docx.internal.content;

import java.awt.Image;
import java.io.ByteArrayInputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.commons.lang.builder.ToStringBuilder;
import org.apache.commons.lang.builder.ToStringStyle;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;

public class ImageContent extends Content {
	private RelationshipsPart relationshipPart;
	private String relationshipId;

	public ImageContent(int width, int height, RelationshipsPart relationshipPart, String relationshipId) {
		super(width, height);
		this.relationshipPart = relationshipPart;
		this.relationshipId = relationshipId;
	}

	public String getRelationshipId() {
		return relationshipId;
	}

	public Image getImage() throws IOException {
		BinaryPart binary = (BinaryPart) relationshipPart.getPart(relationshipId);

		return ImageIO.read(new ByteArrayInputStream(binary.getBytes()));
	}

	@Override
	public String toString() {
		return new ToStringBuilder(this, ToStringStyle.SHORT_PREFIX_STYLE)
			.append("width", getWidth())
			.append("height", getHeight())
			.append("relationshipId", relationshipId)
			.toString();
	}
}
