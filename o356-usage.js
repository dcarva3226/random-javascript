/*============================================================================================================
 @Include: */ include("ssi.utils.js"); /*
 @Script: ssi.o365.usage.js
 @Author: Danny Carvajal
 @Version: 1.0.0 
 @Date: 3/23/2023
 ================================================================================================*/

const cfg = {
	days		: 33,
	daysAgtCt	: 90,
	debug		: "None",
	minsGt		: 0,
	msgCancel	: "The script has been manually cancelled...",
	update		: true	
};

const rpt = {
	err			: "None",  
	qryTime		: 0,
	totUsrs		: 0,
	totUsrsProc	: 0,
	appFndByProg : 0
};

const tbl = {
	cust		: "cust_o365_usage"
};

const apps = {
	excel : "excel.exe",
	word : "winword.exe",
	ppt : "powerpnt.exe",
	mail : "outlook.exe"
};

const OFFICE = ["Office 365 Enterprise", "Office 365 Personal", "Office 365 Plan E3", "Office 365 ProPlus", "Office Professional", "Office Professional Hybrid", 
	"Office Professional Plus", "Office Professional SR-1", "Office Standard", "Office Enterprise", "Office Home and Business", "Office Premium",
	"Office Home and Student", "Office Professional with FrontPage", "Office Professional with Visual FoxPro", "Office Small Business", "Office ProPlus"];
	
const ms365Filter1 = "%Microsoft 365%";
const ms365Filter2 = "%Microsoft Office 365%";

let run = function() {

	let startDate = new java.util.Date();
	startDate.setDate(startDate.getDate() - cfg.days);	
	
	let agtStartDate = new java.util.Date();
	agtStartDate.setDate(agtStartDate.getDate() - cfg.daysAgtCt);		
	
	if (!summaryIsReady()) {
		throw "Summary routines may still be running. Please contact Scalable.";
	}

	jobState.onProgress(1.0, "Removing old records...");
	batchDeleter(tbl.cust, null);	
		
	jobState.onProgress(1.0, "Running query to read all primary users of machines...");		
	let q = Query.selectDistinct(Query.column("c.primary_user", "user"));
	q.from("cmdb_ci_computer", "c");
	q.join("agt_agent", "a", "a.installed_on", "c.id");
	q.where(AND(NE("c.primary_user", null), GT("a.last_contact_date", agtStartDate)));
	
	let qs = getCurrentMillis();
	let users = this.exec.executeL1(q);		
	rpt.qryTime = Math.round((getCurrentMillis() - qs)/60000);	
	rpt.totUsrs = users.length;
		
	jobState.onProgress(1.0, "Begin user loop...");
	for (let i = 0; i < rpt.totUsrs; i++) {

		let user = users[i];	
		
		if (jobHandle.isStopped()) throw cfg.msgCancel;
		
		if ((rpt.totUsrsProc % 20) ==  0) {				
			setProgress(user);
		}		

		let installed = isOfficeInstalled(user);
			
		if (installed) {
		
			let entity = this.mgr.create(tbl.cust);
			entity.set("user", user);
			entity.set("is_office_installed", installed);
		
			let excelUsageType = usageType(apps.excel, user, startDate);
			if (excelUsageType) {
				entity.set("is_excel_used", true);
				if (excelUsageType==2) entity.set("excel_read_only", true);
			}
						
			let wordUsageType = usageType(apps.word, user, startDate);
			if (wordUsageType) {
				entity.set("is_word_used", true);
				if (wordUsageType==2) entity.set("word_read_only", true);
			}
			
			let pptUsageType = usageType(apps.ppt, user, startDate);
			if (pptUsageType) {
				entity.set("is_ppt_used", true);
				if (pptUsageType==2) entity.set("ppt_read_only", true);
			}
			
			let mailUsageType = usageType(apps.mail, user, startDate);
			if (mailUsageType) {
				entity.set("is_outlook_used", true);
				if (mailUsageType==2) entity.set("outlook_read_only", true);
			}
			
			if (cfg.update) entity.save();			
		}

		rpt.totUsrsProc++;
	}
			
	rpt.total++;
};


/* ---------------------------------------------------------------
 Is Office installed on any machine where the user is the 
 primary user.
----------------------------------------------------------------*/
let isOfficeInstalled = function(user) {
	
	let isInstalled = false;
				  
	let criterions = new java.util.ArrayList();
	criterions.add(EQ("c.primary_user", user));
	criterions.add(NE("s.spkg_operational", false));
	criterions.add(IN("sp.id", getMSProducts()));
	
	let q = Query.select(Query.column("s.id", "id"));
	q.from("cmdb_sp_install_summary", "s");
	q.join("cmdb_ci_computer", "c", "c.id", "s.installed_on");
	q.join("cmdb_software_product", "sp", "sp.id", "s.product");
	q.where(AND(criterions));
	q.limit(1);

	let result = this.exec.execute1(q);
	
	// If we don't find the package, maybe there are exe's with no package reference.
	if (!result) {		
		
		criterions = new java.util.ArrayList();
		criterions.add(Criterion.NE("pi.operational", false));
		criterions.add(EQ("c.primary_user", user));
		criterions.add(Criterion.IN("p.file_name", [apps.word, apps.excel, apps.ppt, apps.mail]));
		
		q = Query.select(Query.column("pi.id", "id"));
		q.from("cmdb_program_instance", "pi");
		q.join("cmdb_ci_computer", "c", "c.id", "pi.installed_on");
		q.join("cmdb_program", "p", "p.id", "pi.program");
		q.where(AND(criterions));
		q.limit(1);	
		result = this.exec.execute1(q);
		if (result) rpt.appFndByProg++;
	}	
	
	if (result) isInstalled = true;
	return isInstalled;
};


/* ---------------------------------------------------------------
 The query below has been optimized to run quickly using the 
 limit 1. Faster than SUM on keystrokes. Usage type is:
 null - no usage
 1 - full usage, meaning mins used > 0 and keystrokes > 0
 2 - read only, meaning mins used > 0 and keystrokes = 0
----------------------------------------------------------------*/
let usageType = function(app, user, startDate) {

	// Full use
	let q = Query.select(Query.column("u.user", "user"));
	q.from("cmdb_program_daily_usage", "u");
	q.join("cmdb_program_instance", "pi", "pi.id", "u.program_instance");
	q.join("cmdb_program", "p", "p.id", "pi.program");
	q.where(AND(EQ("u.user", user), 
				GE("u.usage_date", startDate), 
				EQ("p.file_name", app), 
				GT("u.minutes_in_use", cfg.minsGt), 
				GT("u.keystrokes", 0)));
	q.limit(1);	
	
	let result = this.exec.execute1(q);
	if (result) {
		
		return 1
		
	} else {
		
		// Read only
		let q2 = Query.select(Query.column("u.user", "user"));
		q2.from("cmdb_program_daily_usage", "u");
		q2.join("cmdb_program_instance", "pi", "pi.id", "u.program_instance");
		q2.join("cmdb_program", "p", "p.id", "pi.program");		
		q2.where(AND(EQ("u.user", user), 
					GE("u.usage_date", startDate), 
					EQ("p.file_name", app), 
					GT("u.minutes_in_use", cfg.minsGt), 
					EQ("u.keystrokes", 0)));
		q2.limit(1);	
		
		let result2 = this.exec.execute1(q2);
		if (result2) return 2
	}
	
	return null;
};


/* ---------------------------------------------------------------
 Get the product IDs for those products we wish to check for 
 their install statuses.
----------------------------------------------------------------*/
let getMSProducts = function() {
	
	let qp = Query.select(Query.column("sp.id", "id"));
	qp.from("cmdb_software_product", "sp");
	qp.where(
		OR(IN("sp.name", this.OFFICE), 
			Criterion.ILIKE("sp.name", this.ms365Filter1), 
			Criterion.ILIKE("sp.name", this.ms365Filter2))	
	);
	
	return this.exec.executeL1(qp);	
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
	this.cfg.debug += ", log cnt: " + count + "/4 after " + startDate;
	if (count >= 4) pass = true;		
	
	return pass;	
};


/* ---------------------------------------------------------------
 Set job progress
----------------------------------------------------------------*/
let setProgress = function(user) {
	let percentage = ((rpt.totUsrsProc / rpt.totUsrs) * 100.0);
			jobState.onProgress(percentage, 
				String.format("{0} out of {1} users processed. Qry time: {2}, Cur user: {3}, App found by prog inst: {4}...", 
					rpt.totUsrsProc,
					rpt.totUsrs,
					rpt.qryTime,
					user,
					rpt.appFndByProg));	
};


/* ----------------------------------------------------------------------------------------------------------------

 STARTING POINT

---------------------------------------------------------------------------------------------------------------- */ 
try {
  
	if (!jobHandle.isStopped()) {
		run();    
	} else {
		rpt.err = cfg.msgCancel;
	}     
  
} catch (e) {
  
	rpt.err = e;
  
} finally {
  
	let result = String.format("Users: {0} out of {1}, last error: {2}, Qry time: {3}, Apps found thru prog inst: {4}, update flag = {5}, debug={6}", 
		rpt.totUsrsProc,
		rpt.totUsrs,
		rpt.err,
		rpt.qryTime,
		rpt.appFndByProg,
		cfg.update,
		cfg.debug);

	jobState.onProgress(100.0, result);			
};