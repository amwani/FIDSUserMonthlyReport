package com.demo.ExcelProject;

import java.util.Date;

public class Users {
	Integer itemId;
	String userid;
	String username;
	String idDescription;
	String userStatus;
	Date dateOfStatus;
	Date lastSignedOnDate;
	Integer noOfReset;
	String branchCode;
	

	public Users(Integer itemId, String userid, String username, String idDescription, String userStatus, Date dateOfStatus, Date lastSignedOnDate, Integer noOfReset, String branchCode) {
		super();
		this.itemId = itemId;
		this.userid = userid;
		this.username = username;
		this.idDescription = idDescription;
		this.userStatus = userStatus;
		this.dateOfStatus = dateOfStatus;
		this.lastSignedOnDate = lastSignedOnDate;
		this.noOfReset = noOfReset;
		this.branchCode = branchCode;
	}
	public Integer getItemId() {
		return itemId;
	}
	public void setItemId(Integer itemId) {
		this.itemId = itemId;
	}
	public String getUserid() {
		return userid;
	}
	public void setUserid(String userid) {
		this.userid = userid;
	}
	public String getUsername() {
		return username;
	}
	public void setUsername(String username) {
		this.username = username;
	}
	public String getIdDescription() {
		return idDescription;
	}
	public void setIdDescription(String idDescription) {
		this.idDescription = idDescription;
	}
	public String getUserStatus() {
		return userStatus;
	}
	public void setUserStatus(String userStatus) {
		this.userStatus = userStatus;
	}
	public Date getDateOfStatus() {
		return dateOfStatus;
	}
	public void setDateOfStatus(Date dateOfStatus) {
		this.dateOfStatus = dateOfStatus;
	}
	public Date getLastSignedOnDate() {
		return lastSignedOnDate;
	}
	public void setLastSignedOnDate(Date lastSignedOnDate) {
		this.lastSignedOnDate = lastSignedOnDate;
	}
	public Integer getNoOfReset() {
		return noOfReset;
	}
	public void setNoOfReset(Integer noOfReset) {
		this.noOfReset = noOfReset;
	}
	public String getBranchCode() {
		return branchCode;
	}
	public void setBranchCode(String branchCode) {
		this.branchCode = branchCode;
	}

}
