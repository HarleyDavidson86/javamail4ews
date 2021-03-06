/*
The JavaMail4EWS project.
Copyright (C) 2011  Sebastian Just

This library is free software; you can redistribute it and/or
modify it under the terms of the GNU Lesser General Public
License as published by the Free Software Foundation; either
version 3.0 of the License, or (at your option) any later version.

This library is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
Lesser General Public License for more details.

You should have received a copy of the GNU Lesser General Public
License along with this library; if not, write to the Free Software
Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 */
package org.sourceforge.net.javamail4ews.transport;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Enumeration;
import java.util.UUID;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.mail.Address;
import javax.mail.BodyPart;
import javax.mail.Header;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Part;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.URLName;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
import microsoft.exchange.webservices.data.core.enumeration.property.DefaultExtendedPropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.MapiPropertyType;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.property.complex.FileAttachment;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.property.definition.ExtendedPropertyDefinition;

import org.apache.commons.configuration.Configuration;
import org.sourceforge.net.javamail4ews.util.Util;

import com.sun.mail.smtp.SMTPSendFailedException;

public class EwsTransport extends Transport {
	private static final String TEXT_STAR = "text/*";
    private static final String MULTIPART_ALTERNATIVE = "multipart/alternative";
    private static final String MULTIPART_MIXED = "multipart/mixed";
    private static final String MULTIPART_STAR = "multipart/*";
    private static final String TEXT_PLAIN = "text/plain";
	private static final String TEXT_HTML = "text/html";
	
	private static final Logger logger = Logger.getLogger(EwsTransport.class.getName());
	
	private ExchangeService service;
	private String protocol;
	
	public EwsTransport(Session session, URLName urlname) {
		super(session, urlname);
	}

    private String getProtocol() {
        if (protocol == null) 
            protocol = session.getProperty("mail.store.protocol");
        if (protocol == null || protocol.trim().isEmpty())
            protocol = "ewsstore";
        return protocol;
    }
	
	@Override
	protected boolean protocolConnect(String host, int port, String user, String password) throws MessagingException {
        if (user == null)
            user = session.getProperty("mail." + getProtocol() + ".user");
        if (password == null)
            password = session.getProperty("mail." + getProtocol() + ".password");
		service = Util.getExchangeService(getProtocol(), host, port, user, password, session);
		if (service == null) {
			return false;
		} else {
			return true;
		}
	}
	
	public void sendMessage(Message pMessage, Address[] addresses, Address[] ccaddresses, Address[] bccaddresses) throws MessagingException {
		try {
			EmailMessage msg = new EmailMessage(getService());
			createHeaders(msg, pMessage);

			createAddresses(msg, pMessage, addresses, ccaddresses, bccaddresses);
			createSubject(msg, pMessage);
			createBody(msg, pMessage);

			sendMessage(msg);
			
		} catch (MessagingException e) {
			throw e;
		} catch (Exception e) {
		    String message = e.getMessage();
            if (message != null && message.contains(
                    "The user account which was used to submit this request does not have the right to send mail"
                            + " on behalf of the specified sending account")) {
                SMTPSendFailedException ex = new SMTPSendFailedException("send", 551,
                        "Could not send : insufficient right to send on behalf of '" + pMessage.getFrom()[0] + "'", e,
                        null, pMessage.getAllRecipients(), null);
                // (
                // "Could not send : insufficient right to send on behalf of " + pMessage.getFrom()[0], e);
                throw ex;
            } else
                if(message != null)
                    throw new MessagingException(message, e);
                else
                    throw new MessagingException("no detailed message provided",e);
		}
	}
	
	private void sendMessage(EmailMessage msg) throws Exception {
		if (getConfiguration().getBoolean("org.sourceforge.net.javamail4ews.transport.EwsTransport.SendAndSaveCopy")) {
			msg.sendAndSaveCopy(WellKnownFolderName.SentItems);
		} else {
			msg.send();
		}
	}

	private void createHeaders(EmailMessage msg, Message message) throws Exception {
		if(message != null){
			Enumeration<Header> headers = message.getAllHeaders();
			ExtendedPropertyDefinition extendedPropertyDefinition = null;
			while(headers.hasMoreElements()){
				Header header = (Header) headers.nextElement();
				extendedPropertyDefinition = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, header.getName(),MapiPropertyType.String);
				msg.setExtendedProperty(extendedPropertyDefinition,header.getValue());
			}
		}
	}

	public void sendMessage(Message pMessage, Address[] addresses, Address[] ccaddresses) throws MessagingException {
		sendMessage(pMessage,addresses,ccaddresses, new Address[0]);
	}
	
	@Override
	public void sendMessage(Message pMessage, Address[] addresses) throws MessagingException {
		sendMessage(pMessage,addresses, new Address[0]);
	}

	private byte[] bodyPart2ByteArray(BodyPart pPart) throws IOException, MessagingException {
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		pPart.getDataHandler().writeTo(os);
		return os.toByteArray();
	}
	
	private void createBody(EmailMessage msg, Message message) throws Exception {
		MessageBody mb;
		mb = createBodyFromPart(msg,message, false);
		msg.setBody(mb);
	}
	
	private String getFirstHeaderValue(BodyPart part, String pKey) throws MessagingException {
		Enumeration<Header> lMatchingHeaders = part.getMatchingHeaders(new String[]{pKey});
		
		if (lMatchingHeaders.hasMoreElements()) {
			Header lHeader = lMatchingHeaders.nextElement();
			String lValue = lHeader.getValue();
			
			return lValue;
		}
		return null;
	}
    
    private MessageBody createBodyFromPart(EmailMessage msg, Part part, boolean treatAsAttachement)
            throws MessagingException, IOException, ServiceLocalException {

        MessageBody mb = new MessageBody();
        if (part.isMimeType(TEXT_PLAIN)) {
            String s = (String) part.getContent();
            mb.setBodyType(BodyType.Text);
            mb.setText(s);
        } else if (part.isMimeType(TEXT_STAR)) {
            logger.log(Level.FINE, "mime-type is '" + part.getContentType() + "' handling as " + TEXT_HTML);
            String s = (String) part.getContent();
            mb.setBodyType(BodyType.HTML);
            mb.setText(s);
        } else if (part.isMimeType(MULTIPART_ALTERNATIVE) && !treatAsAttachement) {
            logger.log(Level.FINE, "mime-type is '" + part.getContentType() + "'");
            Multipart mp = (Multipart) part.getContent();
            String text1 = "";
            String text2 = "";
            for (int i = 0; i < mp.getCount(); i++) {
                Part p = mp.getBodyPart(i);
                if (p.isMimeType(TEXT_HTML)) {
                    text1 += p.getContent();
                    mb.setText(text1);
                    mb.setBodyType(BodyType.HTML);
                }
                if (p.isMimeType(TEXT_PLAIN)) {
                    text2 += p.getContent();
                    mb.setText(text2);
                    mb.setBodyType(BodyType.Text);
                }
            }
            if (!treatAsAttachement)
                createBodyFromPart(msg, part, true);
        } 
        else if (part.isMimeType(MULTIPART_STAR) && !part.isMimeType(MULTIPART_ALTERNATIVE)) {
            logger.log(Level.FINE, "mime-type is '" + part.getContentType() + "'");
            Multipart mp = (Multipart) part.getContent();
            int start = 0;
            if (!treatAsAttachement) {
                mb = createBodyFromPart(msg, mp.getBodyPart(start), false);
                start++;
            }
            for (int i = start; i < mp.getCount(); i++) {
                BodyPart lBodyPart = mp.getBodyPart(i);
                byte[] lContentBytes = bodyPart2ByteArray(lBodyPart);

                FileAttachment lNewAttachment;

                String lContentId = getFirstHeaderValue(lBodyPart, "Content-ID");
                if (lContentId != null) {
                    lNewAttachment = msg.getAttachments().addFileAttachment(lContentId, lContentBytes);
                    lNewAttachment.setContentId(lContentId);
                    lNewAttachment.setIsInline(true);

                    logger.log(Level.FINE, "Attached {0} bytes as content {1}", new Object[] { lContentBytes.length, lContentId });
                } else {
                    String fileName = lBodyPart.getFileName();
                    fileName = (fileName == null ? "" + i : fileName);
                    lNewAttachment = msg.getAttachments().addFileAttachment(fileName, lContentBytes);
                    lNewAttachment.setIsInline(false);
                    lNewAttachment.setContentType(lBodyPart.getContentType());

                    logger.log(Level.FINE, "Attached {0} bytes as file {1}", new Object[] { lContentBytes.length, fileName });
                    logger.log(Level.FINE, "content type is {0} ", lBodyPart.getContentType());
                }
                lNewAttachment.setIsContactPhoto(false);
            }
        }
        return mb;
    }

    private void createAddresses(EmailMessage pEmailMessage, Message pMessage, Address[] pToAddresses,
            Address[] pCcAddresses, Address[] pBccAddresses) throws Exception {

        if (pMessage instanceof MimeMessage) {
            MimeMessage lMimeMessage = (MimeMessage) pMessage;

            if (pToAddresses.length <= 0) {
                pToAddresses = lMimeMessage.getRecipients(javax.mail.Message.RecipientType.TO);
            }
            if (pCcAddresses.length <= 0) {
                pCcAddresses = lMimeMessage.getRecipients(javax.mail.Message.RecipientType.CC);
            }

            if (pBccAddresses.length <= 0) {
                pBccAddresses = lMimeMessage.getRecipients(javax.mail.Message.RecipientType.BCC);
            }
        }
        
        Address[] from = pMessage.getFrom();
        if(from != null && from.length > 0) {
            pEmailMessage.setFrom(emailAddressFromInternetAddress(from[0]));
        }
        
        for (Address aAddress : pToAddresses) {
            logger.log(Level.INFO, "Adding adress {0} as TO recepient", aAddress.toString());
            pEmailMessage.getToRecipients().add(emailAddressFromInternetAddress(aAddress));
        }
        if (pCcAddresses != null) {
            for (Address aAddress : pCcAddresses) {
                logger.log(Level.INFO, "Adding adress {0} as CC recepient", aAddress.toString());
                pEmailMessage.getCcRecipients().add(emailAddressFromInternetAddress(aAddress));
            }
        }
        if (pBccAddresses != null) {
            for (Address aAddress : pBccAddresses) {
                logger.log(Level.INFO, "Adding adress {0} as BCC recepient", aAddress.toString());
                pEmailMessage.getBccRecipients().add(emailAddressFromInternetAddress(aAddress));
            }
        }
    }

    private EmailAddress emailAddressFromInternetAddress(Address address) {
        String personalPart = "";
        String internetPart = "";
        if (isInternetAddress(address)) {
            personalPart = ((InternetAddress) address).getPersonal();
            internetPart = ((InternetAddress) address).getAddress();
        } else {
            internetPart = address.toString();
        }
        if (personalPart != null && !personalPart.isEmpty()) {
            logger.info("creating address : personal part is '"+personalPart+"' internet part is '"+internetPart+"'");
            return new EmailAddress(personalPart, internetPart);
        }
        logger.info("creating address : internet part is '"+internetPart+"'");
        return new EmailAddress(internetPart);
    }

    private boolean isInternetAddress(Address anAddress){
		
		if (anAddress instanceof InternetAddress) {
			return true;
		}
		return false;
	}
	private void createSubject(EmailMessage msg, Message message) throws MessagingException, Exception {
		msg.setSubject( message.getSubject() );
	}

	private ExchangeService getService() {
		return service;
	}
	
	private Configuration getConfiguration() {
		return Util.getConfiguration(session);
	}
}
