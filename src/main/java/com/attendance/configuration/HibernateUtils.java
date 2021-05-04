package com.attendance.configuration;

import org.hibernate.SessionFactory;
import org.hibernate.cfg.Configuration;

public class HibernateUtils {
	private static SessionFactory sessionFactory;
	
	public static SessionFactory getSessionFactory() {
		return sessionFactory;
	}
	
	static {
		Configuration configuration = new Configuration().configure();
		sessionFactory = configuration.buildSessionFactory();
	}
}

	