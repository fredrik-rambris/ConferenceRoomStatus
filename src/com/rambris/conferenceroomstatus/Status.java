package com.rambris.conferenceroomstatus;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.Collection;
import java.util.Date;

import microsoft.exchange.webservices.data.Appointment;
import microsoft.exchange.webservices.data.CalendarFolder;
import microsoft.exchange.webservices.data.CalendarView;
import microsoft.exchange.webservices.data.EmailAddress;
import microsoft.exchange.webservices.data.ExchangeCredentials;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.FindItemsResults;
import microsoft.exchange.webservices.data.FolderId;
import microsoft.exchange.webservices.data.Mailbox;
import microsoft.exchange.webservices.data.WebCredentials;
import microsoft.exchange.webservices.data.WellKnownFolderName;

public class Status
{
	public String username;
	public String password;
	public String exchangeURI;
	public String roomlist;
	
	private ExchangeService service;
	public Status() throws URISyntaxException
	{
		service = new ExchangeService();
		ExchangeCredentials credentials = new WebCredentials(username, password);
		service.setCredentials(credentials);
		service.setUrl(new URI(exchangeURI));
	}
	
	public Collection<EmailAddress>getRooms(String address) throws Exception
	{
		return service.getRooms(new EmailAddress(address));

	}
	
	public void getRoomStatus(String room) throws Exception
	{
		Mailbox mailbox = new Mailbox(room);
		FolderId calendarFolder = new FolderId(WellKnownFolderName.Calendar, mailbox);
		CalendarFolder calendar=CalendarFolder.bind(service, calendarFolder);
		Date start = new Date();
		Date end = new Date(start.getTime() + (1000 * 3600 * 8));
		CalendarView view = new CalendarView(start, end, 50);

        FindItemsResults<Appointment> appointments = calendar.findAppointments(view);
        
        for(Appointment a:appointments)
        {
        	a.load();
        	System.out.println();
        	System.out.println("Subject: " + a.getSubject());
        	System.out.println("Start: " + a.getStart());
        	System.out.println("End: " + a.getEnd());
        	System.out.println("Location: " + a.getLocation());
        	System.out.println("Organizer: " + a.getOrganizer());
        	System.out.println("Duration: " + a.getDuration());
        	System.out.println("-----------------------------------------");
        }
	
	
	}
	

	public static void main(String[] args) throws Exception
	{
		Status status=new Status();
		status.exchangeURI=args[0];
		status.username=args[1];
		status.password=args[2];
		status.roomlist=args[3];
		
		for(EmailAddress room:status.getRooms(status.roomlist))
		{
			status.getRoomStatus(room.getAddress());
		}
	}

}
