/*============================================================================================================
 @Include: */ include("ssi.utils.js"); /*
 @Script: ssi.report.harvesting.js
 @Author: Danny Carvajal
 @Version: 1.4.0 
 @Date: 4/4/2022
	1.1.0: Blank software install date value should be False it is reflecting as True.
	       Removed VDI-DR from criterion.
	1.2.0: 8/10/2023 - Added support for 90, 30 or 7 day runs.
	1.3.0: 8/23/2023 - Added support for Docker install dates. See getInstallDate().
    1.4.0: 3/11/2025 - Added support for query filtering min agt version, install date age, last full inv
	       date age. Added more pdd.status filters.

 The product to run against is stored in this table: cust_report_harvesting_r19_config. Only one product
 should exist in this table. Usage period for this report is found in the config table.

 Software "InstallDate" - this is the MIN package "install_date" having related software product version.
 
 Staffid is gathered from computer's primary user:
 Staffid = installed_on>primary_user>person>proton_user_data>staffid
 
 Note: getLastDayUsed() is the slowest part of this report. It is based off of days in config table. It is possible 
 to speed up this script by moving the getLastDayUsed() queries into the join. However, this causes the 
 main query to go from < 0 minutes to 20 minutes. On a large instance, such a query could use up a lot 
 of memory.
=============================================================================================================*/

var cfg = {
	days 		: 0,
	debug		: "None",
	dcrDtopProd : "Docker Desktop",
	dcrWinProd  : "Docker CE for Windows",
	dcrWinPrdId : null,
	isDocker	: false,
	update		: true
};

var rpt = {	
	err 		: "None",  
	noproton	: 0,
	product     : "n/a",
	qryTime     : 0,
	total		: 0
};

var tbl = {
	hrvst		: "cust_report_harvesting_r19",
	hrvstCfg	: "cust_report_harvesting_r19_config"
};


var run = function() {
	
	this.cfg.days = getConfig("days");
	const usagePeriod = getUsagePeriod(cfg.days, 1);
	const minAgtVersion = getConfig("minimum_agent_version");
	const installDateAge = getConfig("install_date_age");
	const inventoryDateAge = getConfig("inventory_date_age");
	const product = getConfig("product");	
	if (!product) throw "No product was found in the " + hrvstCfg + " table...";
	this.rpt.product = getProductName(product);
	
	jobState.onProgress(1.0, "Removing old records...");
	batchDeleter(tbl.hrvst, null);
	
	// useful if Docker is the being run against
	if (this.rpt.product == this.cfg.dcrDtopProd) {
		
		this.cfg.isDocker = true;
		this.cfg.dcrWinPrdId = getProductId(this.cfg.dcrWinProd);
	}
	
	if (!summaryDaysUsedReady()) {
		throw "Summary data needs to be reviewed. Please contact Scalable.";
	}
	
	if (!summaryIsReady()) {
		throw "Summary routines may still be running. Please contact Scalable.";
	}
	
	if (this.cfg.days != 90 && this.cfg.days != 30 && this.cfg.days != 7) {
		throw "The reporting days for this report should either be 90, 30 or 7.";		
	}
	
	jobState.onProgress(1.0, "Running query for product: " + this.rpt.product + ", days: " + this.cfg.days + "...");
	
	let criterions = new java.util.ArrayList();
	criterions.add(NE("s.spkg_operational", false));
	criterions.add(EQ("spv.product", product));	
	criterions.add(Criterion.NOT_IN("pdd.comment", ["Robotic", "DR"]));	
	criterions.add(Criterion.NOT_IN("pdd.status", ["Disposed", "Stolen", "Settled", "Returned to Vendor", "Obsolete", "In Transit", "In Storage", "For Disposal-Refurb", "Obsolete", "Sent External", "Discovered", "Divested - HW Only", "Under Investigation", "Divested - HW and Data", "Charity Donation"]));
	if (minAgtVersion) criterions.add(GE("a.package_installed", minAgtVersion));

	// Only show records where inventory > a certain date
	if (inventoryDateAge) {
		let inventoryMinDate = new java.util.Date();
		inventoryMinDate.setDate(inventoryMinDate.getDate() - inventoryDateAge);
		criterions.add(GE("a.last_full_inventory", inventoryMinDate));
	}
	
	let cols = new java.util.ArrayList();
	cols.add(Query.column("s.installed_on", "installed_on"));
	cols.add(Query.column("s.first_tracked_on", "first_tracked_on"));
	cols.add(Query.column("s.software", "software"));
	cols.add(Query.column("a.last_contact_date", "last_contact_date"));
	cols.add(Query.column("pud.staffid", "staffid1"));
	cols.add(Query.column("pud2.staffid", "staffid2"));
	cols.add(Query.column("pud.id", "proton_user_data1"));
	cols.add(Query.column("pud2.id", "proton_user_data2"));
	
	// Get a list of computers with the software products installed	
	let q = Query.select(cols);
	q.from("cmdb_spv_install_summary", "s");
	q.join("cmdb_ci_computer", "c", "c.id", "s.installed_on");
	q.join("cmdb_software_product_version", "spv", "spv.id", "s.software");
	
	// Get staffid by primary user
	q.leftJoin("cmn_user", "u", "u.id", "c.primary_user");
	q.leftJoin("cmn_person", "p", "p.id", "u.person"); 
	q.leftJoin("proton_user_data", "pud", "pud.id", "p.proton_user_data");
	
	// Get staffid by owner (just in case there is no primary user)
	q.leftJoin("cmn_user", "u2", "u2.id", "c.owner");
	q.leftJoin("cmn_person", "p2", "p2.id", "u2.person"); 
	q.leftJoin("proton_user_data", "pud2", "pud2.id", "p2.proton_user_data");
	
	q.leftJoin("proton_device_data", "pdd", "pdd.id", "c.proton_device_data");
	q.leftJoin("agt_agent", "a", "a.installed_on", "s.installed_on");
	q.orderBy("pud.staffid", Order.ASC); // do not remove
	q.orderBy("pud2.staffid", Order.ASC); // do not remove	
	q.orderBy("s.software", Order.ASC); // do not remove
	q.orderBy("s.first_tracked_on", Order.ASC); // do not remove
	q.where(AND(criterions));
	
	let qs = getCurrentMillis();
	let spvs = exec.executeLM(q);
	rpt.qryTime = Math.round((getCurrentMillis() - qs)/60000);
	let lastStaffId = null;
	
	for (let i = 0; i < spvs.length; i++) {        

		let usedBySid = false;

		if ((rpt.total % 10) ==  0) {
			if (jobHandle.isStopped()) throw "Script job was cancelled...";
		}			
			 
		let spv = spvs[i];
		let installedOn = spv["installed_on"];
		let protonUserFromPrimary = spv["proton_user_data1"];		
		let protonUserFromOwner = spv["proton_user_data2"];		
		let protonUser = (protonUserFromPrimary != null) ? protonUserFromPrimary : protonUserFromOwner;
		let software = spv["software"];
		let staffid1 = spv["staffid1"];
		let staffid2 = spv["staffid2"];
		let staffid = (staffid1 != null) ? staffid1 : staffid2;
		let computerCount = (staffid) ? getComputerCount(software, staffid) : 0;
		let lastContact = spv["last_contact_date"];
		let firstTrackedOn = spv["first_tracked_on"];
		let startDate = new java.util.Date();
		startDate.setDate(startDate.getDate() - cfg.days);		
		let installDate = getInstallDate(software, installedOn);
		let daysUsed = getSummaryDaysUsed(installedOn, usagePeriod, software);
		let lastDayUsed = (daysUsed > 0) ? getLastDayUsed(installedOn, software, startDate) : null;
		let version = getVersion(software, installedOn);
		let installSource = getInstallSource(software, installedOn);
		
		// Make sure install date is within installDateAge
		if (installDateAge) {
			if (installDateAge < this.getDateDiff(installDate, new java.util.Date())) {
				continue;
			}
		}

		// Found cases where the spv install summary is duplicated based on first tracked on.
		// This readEntity check is faster than grouping on First Tracked On.
		let crit1 = EQ("installed_on", installedOn);
		let crit2 = EQ("software", software);
		
		let entity = mgr.readEntity(tbl.hrvst, AND(crit1, crit2));
		if (!entity) {
			entity = mgr.create(tbl.hrvst);
			entity.set("software", software);
			entity.set("installed_on", installedOn);
			entity.set("days_used", daysUsed);
			entity.set("used_in_period", (daysUsed > 0) ? true : false); 
			entity.set("installed_in_period", (getDateDiff(installDate, new Date()) <= this.cfg.days || installDate == null) ? true : false); 
			entity.set("first_tracked_in_period", (getDateDiff(firstTrackedOn, new Date()) <= this.cfg.days) ? true : false); 
			entity.set("first_tracked_on", firstTrackedOn);
			entity.set("computer_count", (computerCount > 0) ? computerCount : 1);
			entity.set("last_contact_date", lastContact);
			entity.set("last_contact_flag", (lastContact != null && getDateDiff(lastContact, new java.util.Date()) <= 7) ? true : false);
			entity.set("last_day_used", lastDayUsed);
			entity.set("install_date", installDate);
			entity.set("version", version);
			entity.set("install_source", installSource);
			entity.set("staffid", staffid); // Need this stored for use by setAnyMachineUsedFlag()
			
			if (!protonUser) {
				rpt.noproton++;			
			}
			
			if (cfg.update) entity.save();
		}

		// This staffid had at least one machine used?
		if (daysUsed > 0) usedBySid = true;

		// Do we need to update staffid records to show that at least one machine was used?
		if (lastStaffId != staffid) {
			setAnyMachineUsedFlag(lastStaffId);
		}		
		
		// Handle last record queried
		if (i == spvs.length-1) {
			setAnyMachineUsedFlag(staffid);				
		}
		
		lastStaffId = staffid;
		rpt.total++;
		
		if ((this.rpt.total % 25) ==  0) {				
			let percentage = ((this.rpt.total / spvs.length) * 100.0);
			jobState.onProgress(percentage, 
				String.format("{0} out of {1} records processed. Query time {2} mins...", 
					this.rpt.total,
					spvs.length,
					(this.rpt.qryTime > 0) ? " = " + this.rpt.qryTime : " < " + this.rpt.qryTime));
		}
	}

	// Reset certain config values just in case user forget to clear them out on next run
	resetConfigValues();
};


/* ------------------------------------------------------------------------
  We need computer counts grouped by staff id. Different
  devices can have the same staffid. So this query needs
  to be run separately and a subquery won't work here
  due to staffid just being a string value. (not a reference)
 ------------------------------------------------------------------------ */
let getComputerCount = function(software, staffid) {

	let criterion = new java.util.ArrayList();
	criterion.add(EQ("s.software", software));
	criterion.add(NE("s.spkg_operational", false));
	criterion.add(OR(
			EQ("pud.staffid", staffid), 
			AND(EQ("pud2.staffid", staffid), EQ("c.primary_user", null))));
	
	let col = Query.coalesce(Query.countDistinct("s.installed_on"), Query.value(0)).as("computer_count");
	let q = Query.select(java.util.Arrays.asList(col));
	q.from("cmdb_spv_install_summary", "s");
	q.join("cmdb_ci_computer", "c", "c.id", "s.installed_on");	

	q.leftJoin("cmn_user", "u", "u.id", "c.primary_user");
	q.leftJoin("cmn_person", "p", "p.id", "u.person");
	q.leftJoin("proton_user_data", "pud", "pud.id", "p.proton_user_data");

	// No match on primary user? Try owner.
	q.leftJoin("cmn_user", "u2", "u2.id", "c.owner");
	q.leftJoin("cmn_person", "p2", "p2.id", "u2.person");
	q.leftJoin("proton_user_data", "pud2", "pud2.id", "p2.proton_user_data");

	q.where(AND(criterion));	
	return this.exec.execute1(q);
};


let getLastDayUsed = function(installedOn, software, startDate) {
		
	let criterions = new java.util.ArrayList();
	criterions.add(EQ("spkg.software", software));
	criterions.add(GE("du.usage_date", startDate));
	criterions.add(EQ("du.used_from", installedOn));
	criterions.add(GT("du.minutes_in_use", 0));
	
	let list = new java.util.ArrayList();
	list.add(Query.max("du.usage_date").as("last_day_used"));
	
	let q = Query.select(list);
	q.from("cmdb_program_daily_usage", "du");
	q.join("cmdb_program_instance", "pi", "pi.id", "du.program_instance");
	q.join("cmdb_ci_spkg", "spkg", "spkg.id", "pi.spkg");
	q.where(AND(criterions));
	return this.exec.execute1(q);
};


let getSummaryDaysUsed = function(installedOn, usagePeriod, software) {
	
	let criterion = new java.util.ArrayList();
	criterion.add(EQ("used_from", installedOn));
	criterion.add(EQ("period", usagePeriod));
	criterion.add(GT("minutes_in_use", 0));
	criterion.add(EQ("software", software));
	
	let q = Query.select(Query.sum("u.days_used").as("days_used"));
	q.from("cmdb_spv_usage_summary", "u");
	q.where(AND(criterion));	
	let daysUsed = this.exec.execute1(q);	
	return (daysUsed > 0) ? daysUsed : 0;
};


/* ------------------------------------------------------------------------
  Not getting installDate this via subquery as we have enough joins already. 
  Plus Docker requires special handling. If current SPV is Docker, we need 
  to perform a special query to get install date from a particular package.
 ------------------------------------------------------------------------ */
let getInstallDate = function(spv, installedOn) {
	
	let criterions = new java.util.ArrayList();
	
	if (!this.cfg.isDocker) {
		criterions.add(EQ("spkg.software", spv));
	} else {
		criterions.add(EQ("spv.product", this.cfg.dcrWinPrdId));
	}
	
	criterions.add(EQ("spkg.installed_on", installedOn));
	criterions.add(NE("spkg.operational", false));	
	
	let list = new java.util.ArrayList();
	list.add(Query.min("spkg.install_date").as("install_date"));
		
	let q = Query.select(list);
	q.from("cmdb_ci_spkg", "spkg");
	if (this.cfg.isDocker) q.join("cmdb_software_product_version", "spv", "spv.id", "spkg.software");
	q.where(AND(criterions));
	return this.exec.execute1(q);
};


/* ------------------------------------------------------------------------
  Grab package versions associated with the SPV. There could be > 1.
 ------------------------------------------------------------------------ */
let getVersion = function(spv, installedOn) {
	
	let ret = null;
	let criterions = new java.util.ArrayList();
	criterions.add(EQ("spkg.software", spv));
	criterions.add(EQ("spkg.installed_on", installedOn));
	criterions.add(NE("spkg.operational", false));	
	
	let list = new java.util.ArrayList();
	list.add(Query.min("spkg.version").as("version"));	
	
	let q = Query.selectDistinct(list);
	q.from("cmdb_ci_spkg", "spkg");
	q.where(AND(criterions));
	let versions = this.exec.executeL1(q);
	
	if (versions) {
		if (versions.length > 1) {
			for (let i = 0; i < versions.length; i++) {
				ret += versions[i] + ",";
			}
		} else {
			ret = versions[0];
		}
	}
	
	return ret;
};


/* ------------------------------------------------------------------------
  Grab install sources associated with the SPV. There could be > 1.
 ------------------------------------------------------------------------ */
let getInstallSource = function(spv, installedOn) {
	
	let ret = null;
	let criterions = new java.util.ArrayList();
	criterions.add(EQ("spkg.software", spv));
	criterions.add(EQ("spkg.installed_on", installedOn));
	criterions.add(NE("spkg.operational", false));	
	
	let list = new java.util.ArrayList();
	list.add(Query.min("spkg.install_source").as("install_source"));	
	
	let q = Query.selectDistinct(list);
	q.from("cmdb_ci_spkg_windows", "spkg");
	q.where(AND(criterions));
	let ins = this.exec.executeL1(q);
	
	if (ins) {
		if (ins.length > 1) {
			for (let i = 0; i < ins.length; i++) {
				ret += ins[i] + ",";
			}
		} else {
			ret = ins[0];
		}
	}
	
	if (ret) {
		if (ret.length > 500) ret = ret.substring(0, 500);
	}
	return ret;	
};


/* ------------------------------------------------------------------------
  This is used to set the any_machine_used_in_period column to true or 
  false if it was used by a given staffid.
 ------------------------------------------------------------------------ */
let setAnyMachineUsedFlag = function(staffid) {
	
	if (staffid == null) return;
	var wasAnyMachineUsedByStaffId = false;
	
	let criterions = new java.util.ArrayList();
	criterions.add(EQ("staffid", staffid));
	criterions.add(GT("days_used", 0));
	
	// Were any machine used for this staffid?
	let cnt = this.mgr.query(tbl.hrvst).where(Criterion.AND(criterions)).count();	
	if (cnt > 0)
		wasAnyMachineUsedByStaffId = true;
	else
		wasAnyMachineUsedByStaffId = false;
	
	let batchUpdate = dbApi.createBatchUpdate(tbl.hrvst);
	batchUpdate.set("any_machine_used_in_period", wasAnyMachineUsedByStaffId);
	batchUpdate.update(EQ("staffid", staffid));	
};


/* ------------------------------------------------------------------------
  Read the confg value in the Harvest config table.
 ------------------------------------------------------------------------ */
let getConfig = function(field) {
	
	// There should only be one record in this table.
	let val = null;
	let crit = NE("id", null);
	
	try {
		val = this.mgr.readEntity(tbl.hrvstCfg, crit).get(field);
	} catch (e if isDatabaseException(e.javaException)) {
		throw "This script only supports one record in the " + tbl.hrvstCfg + " table.";
	}
	
	return val;
};


/* ------------------------------------------------------------------------
  We need to be sure that there is no empty days_used in the summary. This
  script used cmdb_spv_usage_summary.days_used.
 ------------------------------------------------------------------------ */
let summaryDaysUsedReady = function() {
	
	let pass = false;
	
	var crit = EQ("t.days_used", null);
	
	// Make sure the table is not empty
	count = getCount("cmdb_spv_install_summary", null, 10);
	//this.cfg.debug = "spv inst cnt: " + count;
	
	if (count == 10) {
		
		count = getCount("cmdb_spv_usage_summary", null, 10);
		//this.cfg.debug += ", spv usg cnt: " + count;
		
		if (count == 10) { 
			
			// Make sure none have empty Days Used
			count = getCount("cmdb_spv_usage_summary", crit, 1);
			if (count == 0) pass = true;
		}
	}
	
	return pass;
};


/* ------------------------------------------------------------------------
  Go back a week and make sure there isn't a job still running or one that 
  has errored out. If this function runs on a Friday, the startDate is 
  before last Sunday. (7 days back) Thus, criteria should target dates 
  greater than Sunday because that's when summary should have completed.
 ------------------------------------------------------------------------ */
let summaryIsReady = function() {
	
	let pass = false;
	let startDate = new java.util.Date();
	startDate.setDate(startDate.getDate() - 7);		
	
	let criterions = new java.util.ArrayList();	
	criterions.add(GT("end_time", startDate));
	criterions.add(NE("end_time", null));
	criterions.add(EQ("error", null));
	
	let count = mgr.query("cmdb_summary_sql_job_log").where(AND(criterions)).count();
	//this.cfg.debug += ", log cnt: " + count + "/4 after " + startDate;
	if (count >= 4) pass = true;		
	
	return pass;	
};


let getCount = function(table, crit, limit) {
	
	let col = Query.coalesce(Query.column("id")).as("count");
	let q = Query.select(java.util.Arrays.asList(col));
	q.from(table, "t");
	if (crit) q.where(crit);
	q.limit(limit);
	return this.exec.executeL1(q).length;
};

/* ------------------------------------------------------------------------
  Reset certain config values just in case user forget to clear them out 
  on next run. IOW - there are some custom config values that the user
  can use to further filter out records. If the user forgets to clear 
  those out, they will get less data in queries.
 ------------------------------------------------------------------------ */
let resetConfigValues = function() {
	let entity = this.mgr.readEntity(tbl.hrvstCfg, NE("id", null));
	entity.set("minimum_agent_version", null);
	entity.set("install_date_age", null);
	entity.set("inventory_date_age", null);
	entity.save();
};

/* ----------------------------------------------------------------------------------------------------------------

 STARTING POINT

---------------------------------------------------------------------------------------------------------------- */ 
try {
  
	if (!jobHandle.isStopped()) {
		run();    
	} else {
		this.rpt.err = "Script was cancelled...";
	}     
  
} catch (e) {
  
	this.rpt.err = e;
  
} finally {
  
	let result = String.format("{0} recs processed: {1}, no proton matches: {2}, main query time: {3} mins, days: {4}, last error: {5}, update flag = {6}, debug={7}", 
		this.rpt.product,
		this.rpt.total,
		this.rpt.noproton,
		this.rpt.qryTime,
		this.cfg.days,
		this.rpt.err,
		this.cfg.update,
		this.cfg.debug);

	jobState.onProgress(100.0, result);			
};