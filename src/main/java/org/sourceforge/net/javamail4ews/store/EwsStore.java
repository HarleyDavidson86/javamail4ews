/*
 * The JavaMail4EWS project. Copyright (C) 2011 Sebastian Just
 * 
 * This library is free software; you can redistribute it and/or modify it under
 * the terms of the GNU Lesser General Public License as published by the Free
 * Software Foundation; either version 3.0 of the License, or (at your option)
 * any later version.
 * 
 * This library is distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the GNU Lesser General Public License for more
 * details.
 * 
 * You should have received a copy of the GNU Lesser General Public License
 * along with this library; if not, write to the Free Software Foundation, Inc.,
 * 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA
 */
package org.sourceforge.net.javamail4ews.store;

import java.net.MalformedURLException;
import java.net.URL;
import java.util.Optional;
import java.util.regex.Pattern;

import javax.mail.Folder;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.URLName;

import org.apache.commons.configuration.Configuration;
import org.sourceforge.net.javamail4ews.util.Util;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.property.complex.Mailbox;

public class EwsStore extends Store
{

	private ExchangeService service;
	private EwsFolder defaultFolder;
	private String protocol;
	private String sharedAccountAddress;

	public EwsStore(Session session, URLName urlname)
	{
		super(session, urlname);
	}

	private String getProtocol()
	{
		if (protocol == null)
			protocol = session.getProperty("mail.store.protocol");
		if (protocol == null || protocol.trim().isEmpty())
			protocol = "ewsstore";
		return protocol;
	}

	@Override
	protected boolean protocolConnect(String host, int port, String user, String password) throws MessagingException
	{
		if (user == null)
			user = session.getProperty("mail." + getProtocol() + ".user");
		if (password == null)
			password = session.getProperty("mail." + getProtocol() + ".password");
		service = Util.getExchangeService(getProtocol(), host, port, user, password, session);
		return service != null;
	}

	protected Optional<WellKnownFolderName> getWellKnownFolderName(String name)
	{
		if (name == null)
			return Optional.of(WellKnownFolderName.Root);
		name = name.trim();
		for (WellKnownFolderName id : WellKnownFolderName.values())
		{
			if (id.toString().equalsIgnoreCase(name))
			{
				return Optional.of(id);
			}
		}
		return Optional.empty();
	}

	@Override
	public EwsFolder getDefaultFolder() throws MessagingException
	{
		if (defaultFolder == null)
		{
			FolderId folderToAccess = new FolderId(WellKnownFolderName.MsgFolderRoot);
			if (sharedAccountAddress != null)
			{
				folderToAccess = new FolderId(WellKnownFolderName.Inbox, new Mailbox(sharedAccountAddress));
			}
			defaultFolder = new EwsFolder(this, folderToAccess, sharedAccountAddress);
		}
		return defaultFolder;
	}

	@Override
	public EwsFolder getFolder(String name) throws MessagingException
	{
		String[] folderNames = name.split("/");
		Optional<WellKnownFolderName> wellKnownFolderName = getWellKnownFolderName(folderNames[0]);
		EwsFolder rootFolder = null;
		if (wellKnownFolderName.isPresent())
		{
			System.out.println(String.format("Opening WellKnownFolderName matching %s", folderNames[0]));
			if (sharedAccountAddress != null)
			{
				rootFolder = new EwsFolder(this,
						new FolderId(wellKnownFolderName.get(), new Mailbox(sharedAccountAddress)),
						sharedAccountAddress);
			}
			else
			{
				rootFolder = new EwsFolder(this, new FolderId(wellKnownFolderName.get()), sharedAccountAddress);
			}
		}
		try
		{
			if (rootFolder == null)
			{
				rootFolder = getDefaultFolder().getFolder(folderNames[0]);
			}
		}
		catch (Exception e)
		{
			throw new MessagingException(e.getMessage(), e);
		}

		if (rootFolder == null)
		{
			throw new MessagingException("RootFolder '" + folderNames[0] + "' not found.");
		}

		// Unterverzeichnisse finden
		EwsFolder result = rootFolder;
		for (int i = 1; i < folderNames.length; i++)
		{
			String fname = folderNames[i];
			result = result.getFolder(fname);
		}
		return result;

	}

	@Override
	public EwsFolder getFolder(URLName paramUrl) throws MessagingException
	{
		if (paramUrl == null)
			return getDefaultFolder();
		try
		{
			URL url = paramUrl.getURL();
			if (url == null)
				return getDefaultFolder();
			String path = url.getPath();
			if (path == null)
				return getDefaultFolder();
			if (path.startsWith("/"))
				path = path.substring(1);
			EwsFolder result = getDefaultFolder();
			String[] parts = path.split(Pattern.quote("/"));
			for (String part : parts)
			{
				if (part == null || part.length() == 0)
					continue;
				result = result.getFolder(part);
			}
			if (result == null)
				throw new MessagingException("Folder not found");
			return result;
		}
		catch (MalformedURLException e)
		{
			MessagingException err = new MessagingException(e.getMessage(), e);
			err.setStackTrace(e.getStackTrace());
			throw err;
		}
	}

	protected ExchangeService getService()
	{
		return service;
	}

	// Make visible
	@Override
	protected void notifyFolderListeners(int type, Folder folder)
	{
		super.notifyFolderListeners(type, folder);
	}

	@Override
	// Make visible
	protected void notifyConnectionListeners(int type)
	{
		super.notifyConnectionListeners(type);
	}

	protected Configuration getConfiguration()
	{
		return Util.getConfiguration(session);
	}

	/**
	 * @param sharedAccountAddress the sharedAccountAddress to set
	 */
	public void setSharedAccountAddress(String sharedAccountAddress)
	{
		this.sharedAccountAddress = sharedAccountAddress;
	}
}
