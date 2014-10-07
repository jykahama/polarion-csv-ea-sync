using System;
using System.Windows.Forms;
using System.IO;
using LumenWorks.Framework.IO.Csv;
using System.Collections.Generic;
using MicrogoldWindows;

namespace CS_AddinFramework
{
	public class Main
	{
		private bool m_ShowFullMenus = false;
		public Form1 theForm;
        private EA.Repository m_eaRepo;

		//Called Before EA starts to check Add-In Exists
		public String EA_Connect(EA.Repository Repository)
		{
            m_eaRepo = Repository;
			//No special processing required.
			return "a string";
		}

		//Called when user Click Add-Ins Menu item from within EA.
		//Populates the Menu with our desired selections.
		public object EA_GetMenuItems(EA.Repository Repository, string Location, string MenuName) 
		{
            m_eaRepo = Repository;
            EA.Package aPackage = Repository.GetTreeSelectedPackage();
			switch( MenuName )
			{
				case "":
					return "-&Polarion CSV Import";
                case "-&Polarion CSV Import":
					string[] ar = { "&Menu1", "&Menu2", "&Menu3", "About..." };
					return ar;
			}
			return "";
		}
		
		//Sets the state of the menu depending if there is an active project or not
		bool IsProjectOpen(EA.Repository Repository)
		{
            m_eaRepo = Repository;
            try
			{
				EA.Collection c = Repository.Models;
				return true;
			}
			catch
			{
				return false;
			}
		}

		//Called once Menu has been opened to see what menu items are active.
		public void EA_GetMenuState(EA.Repository Repository, string Location, string MenuName, string ItemName, ref bool IsEnabled, ref bool IsChecked)
		{
            m_eaRepo = Repository;
            if (IsProjectOpen(Repository))
            {
                if (Location == "TreeView")
                {
                    IsEnabled = true;
                }
                else
                {
                    if (ItemName == "&Menu1")
                        IsChecked = m_ShowFullMenus;
                    else if (ItemName == "&Menu2")
                        IsEnabled = m_ShowFullMenus;
                    else if (ItemName == "&Menu3")
                        IsEnabled = m_ShowFullMenus;
                }
            }
            else
            {
                // If no open project, disable all menu options
                if (ItemName == "About...")
                    IsEnabled = true;
                else
                    IsEnabled = false;
            }
		}

        private EA.Element CreateChildElement(CsvReader csv, EA.Package parent, String polarionID, Dictionary<String, EA.Element> map)
        {
            if (map.ContainsKey(polarionID))
            {

                EA.Element req = map[polarionID];

                req.Name = csv["Title"];
                req.Notes = csv["Description"] + "\r\n\r\n\r\nMotivation\r\n" + csv["Motivation"] + "\r\n\r\n\r\nNotes\r\n" + csv["Note"];
                // No need to touch polarion ID

                req.PackageID = parent.PackageID;
                req.ParentID = 0; // Directly owned by this Package
                // TODO: Need to fiddle with parents Elements list?

                req.Update();

                map.Remove(polarionID);

                return req;
            }
            else
            {
                EA.Element newReq = (EA.Element)parent.Elements.AddNew(csv["Title"], "Requirement");
                newReq.Notes = csv["Description"] + "\r\n\r\n\r\nMotivation\r\n" + csv["Motivation"] + "\r\n\r\n\r\nNotes\r\n" + csv["Note"];

                EA.TaggedValue polID = (EA.TaggedValue)newReq.TaggedValues.AddNew("PolarionID", "String");
                polID.Value = polarionID;
                polID.Update();

                newReq.Update();

                return newReq;
            }
        }

        private EA.Element CreateChildElement(CsvReader csv, EA.Element parent, String polarionID, Dictionary<String, EA.Element> map)
        {
            if (map.ContainsKey(polarionID))
            {

                EA.Element req = map[polarionID];

                req.Name = csv["Title"];
                req.Notes = csv["Description"] + "\r\n\r\n\r\nMotivation\r\n" + csv["Motivation"] + "\r\n\r\n\r\nNotes\r\n" + csv["Note"];
                // No need to touch polarion ID

                req.PackageID = parent.PackageID; // owned by this Parents Package
                req.ParentID = parent.ElementID;
                // TODO: Need to fiddle with parents Elements list?

                req.Update();

                map.Remove(polarionID);

                return req;
            }
            else
            {
                EA.Element newReq = (EA.Element)parent.Elements.AddNew(csv["Title"], "Requirement");
                newReq.Notes = csv["Description"] + "\r\n\r\n\r\nMotivation\r\n" + csv["Motivation"] + "\r\n\r\n\r\nNotes\r\n" + csv["Note"];

                EA.TaggedValue polID = (EA.TaggedValue)newReq.TaggedValues.AddNew("PolarionID", "String");
                polID.Value = polarionID;
                polID.Update();

                newReq.Update();

                return newReq;
            }
        }

        private String collectPolarionToGuidMap(EA.Collection collection, Dictionary<String, EA.Element> map, String typeOfColl)
        {
            String msg = "";
            if ( typeOfColl == "pkg")
            { 
                foreach (EA.Package p in collection)
                {
                    msg = msg + "Package: " + p.PackageID + p.Name + "\r\n";
                    msg = msg + collectPolarionToGuidMap(p.Packages, map, "pkg");
                    msg = msg + collectPolarionToGuidMap(p.Elements, map, "ele");
                }
            }
            else if ( typeOfColl == "ele")
            { 
                foreach (EA.Element e in collection)
                {
                    msg = msg + "Element: " + e.ElementID + e.Name + "   parentID=" + e.ParentID + "     packageID=" + e.PackageID + "\r\n";
                    msg = msg + collectPolarionToGuidMap(e.Elements, map, "ele");
                    foreach (EA.TaggedValue tv in e.TaggedValues)
                    {
                        if (tv.Name == "PolarionID")
                        {
                            map.Add(tv.Value, e);
                        }
                    }
                }
            }
            return msg;

        }
        //Called when user makes a selection in the menu.
		//This is your main exit point to the rest of your Add-in
		public void EA_MenuClick(EA.Repository Repository, string Location, string MenuName, string ItemName)
		{
            m_eaRepo = Repository;
            switch (ItemName)
			{
				case "&Menu1":
                    {
                        String writerString = "";
                        EA.Package aPackage = Repository.GetTreeSelectedPackage();
                        foreach (EA.Package thePackage in aPackage.Packages)
                        {
                            writerString = writerString + thePackage.Name.ToString() + "," + thePackage.ObjectType.ToString() + "\n";
                        }
                        MessageBox.Show(writerString);
                    }
                    break;
				case "&Menu2":
                    // Read sample data from CSV file
                    using (CsvReader csv = new CsvReader(new StreamReader("C:\\Users\\trehamaljy\\Downloads\\req.csv"), true, ';', '"'))
                    {
                        EA.Package aPackage = Repository.GetTreeSelectedPackage();

                        Dictionary<String, EA.Element> polToGuid = new Dictionary<String, EA.Element>();

                        this.collectPolarionToGuidMap(aPackage.Packages, polToGuid, "pkg");
                        this.collectPolarionToGuidMap(aPackage.Elements, polToGuid, "ele");
                        
                        
                        EA.Element level1 = aPackage.Element;
                        EA.Element level2 = aPackage.Element;
                        EA.Element level3 = aPackage.Element;
                        
                        //String writerString = "";
                        int fieldCount = csv.FieldCount;

                        string[] headers = csv.GetFieldHeaders();
                        //writerString = writerString + String.Join(" ### ", headers) + "\n";
                        while (csv.ReadNextRecord())
                        {
                            
                            if ( csv["ID1"] != "" )
                            {
                                level1 = this.CreateChildElement(csv, aPackage, csv["ID1"], polToGuid);
                            }
                            else if ( csv["ID2"] != "" )
                            {
                                level2 = this.CreateChildElement(csv, level1, csv["ID2"], polToGuid);
                            }
                            else if ( csv["ID3"] != "" )
                            {
                                level3 = this.CreateChildElement(csv, level2, csv["ID3"], polToGuid);
                            }
                            else if ( csv["ID4"] != "" )
                            {
                                this.CreateChildElement(csv, level3, csv["ID4"], polToGuid);
                            }

                        }

                        // Clean up requirements that did not have any references anymore
                        foreach (KeyValuePair<string, EA.Element> kvp in polToGuid)
                        {
                            EA.Element elem = kvp.Value;

                            if (elem.ParentID == 0)
                            {
                                EA.Package parent = Repository.GetPackageByID(elem.PackageID);
                                MessageBox.Show("Deleting " + elem.Name + " from package " + parent.Name);
                                for (short i = 0; i < parent.Elements.Count; i++)
                                {
                                    EA.Element elem_i = (EA.Element)parent.Elements.GetAt(i);
                                    if (elem_i.ElementID == elem.ElementID)
                                    {
                                        MessageBox.Show("Match found");
                                        parent.Elements.DeleteAt(i, true);
                                        break;
                                    }
                                }
                                parent.Update();
                                parent.Elements.Refresh();
                            }
                            else
                            {
                                EA.Element parent = Repository.GetElementByID(elem.ParentID);
                                MessageBox.Show("Deleting " + elem.Name + " from element " + parent.Name);
                                for (short i = 0; i < parent.Elements.Count; i++)
                                {
                                    EA.Element elem_i = (EA.Element)parent.Elements.GetAt(i);
                                    if (elem_i.ElementID == elem.ElementID)
                                    {
                                        MessageBox.Show("Match found");
                                        parent.Elements.DeleteAt(i, true);
                                        break;
                                    }
                                }
                                parent.Update();
                                parent.Elements.Refresh();
                            }

                        }

                        aPackage.Elements.Refresh();
                        //MessageBox.Show(writerString);
                    }
                    break;
                case "&Menu3":
                    {
                        ScrollableMessageBox box = new ScrollableMessageBox();
                        EA.Package aPackage = Repository.GetTreeSelectedPackage();

                        Dictionary<String, EA.Element> polToGuid = new Dictionary<String, EA.Element>();

                        String msg = "";
                        msg = msg + this.collectPolarionToGuidMap(aPackage.Packages, polToGuid, "pkg");
                        msg = msg + this.collectPolarionToGuidMap(aPackage.Elements, polToGuid, "ele");
                        box.Show(msg, "Map");

                        msg = "";
                        foreach (KeyValuePair<string, EA.Element> kvp in polToGuid)
                        {
                            EA.Element elem = kvp.Value;
                            msg = msg + kvp.Key + " -> " + elem.Name + "\r\n";
                        }
                        box.Show(msg, "Map");
                    }
                    break;
                case "About...":
					Form1 anAbout = new Form1();
					anAbout.ShowDialog();					
					break;
			}
		}
	}
}
		
		
	

