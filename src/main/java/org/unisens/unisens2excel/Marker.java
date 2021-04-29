package org.unisens.unisens2excel;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.unisens.Event;

import com.fasterxml.jackson.annotation.JsonFormat;

public class Marker {
	
	final static String DATE_PATTERN = "yyyy-MM-dd'T'HH:mm:ss.SSS";

	@JsonFormat(shape = JsonFormat.Shape.STRING, pattern = DATE_PATTERN, timezone=JsonFormat.DEFAULT_TIMEZONE)
	Date date;
	String type;
	String comment;
	
	public Marker (Date dateTime, Event event, double fs)
	{
		this.date = new Date(dateTime.getTime() + (long)((double)(event.getSampleStamp()/fs) * 1000L));
		this.type = event.getType();
		this.comment = event.getComment();
	}
	
	public Date getDate() {
		return date;
	}

	public String getType() {
		return type;
	}

	public String getComment() {
		return comment;
	}
	
	public String toString()
	{
		return String.format("(%s; %s; %s)", new SimpleDateFormat(DATE_PATTERN).format(date), type, comment);
	}

}
