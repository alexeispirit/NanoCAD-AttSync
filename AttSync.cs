using HostMgd.ApplicationServices;
using HostMgd.EditorInput;
using HostMgd.Runtime;
using Teigha.DatabaseServices;
using Teigha.Runtime;

public class AttSync
{
    RXClass attDefClass = RXClass.GetClass(typeof(AttributeDefinition));

    [LispFunction("att-sync")]
    public void AttSyncFuncBlockRef(ResultBuffer rbArgs)
    {
        // rbArgs - ObjectId of BlockReference (car (entget))
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;
        Editor ed = doc.Editor;

        if (rbArgs == null || rbArgs.AsArray().Length < 0)
        {
            ed.WriteMessage("Wrong argument number");
            return;
        }

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            try
            {
                TypedValue entity = rbArgs.AsArray().First();

                if (entity.TypeCode != (int)LispDataType.ObjectId)
                {
                    throw new System.Exception($"Wrong argument type: {entity.Value.ToString()}.");
                }

                ObjectId id = (ObjectId)entity.Value;
                Entity ent = (Entity)tr.GetObject(id, OpenMode.ForRead);

                string objectName = ent.GetRXClass().Name;
                if (objectName != "AcDbBlockReference")
                {
                    throw new System.Exception($"Wrong argument type: {objectName}.");
                }

                BlockReference br = (BlockReference)ent;

                SynchronizeAttributes(br, tr);
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(ex.Message);
                tr.Abort();
            }
        }
    }

    [LispFunction("att-sync-all")]
    public void AttSyncFuncBlockDef(ResultBuffer rbArgs)
    {
        // rbArgs - ObjectId of BlockDefinition (vlax-vla-object (vla-item (vla-get-blocks (vla-get-activedocument (vlax-get-acad-object))) "BLOCKNAME"))
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;
        Editor ed = doc.Editor;

        if (rbArgs == null || rbArgs.AsArray().Length < 0)
        {
            ed.WriteMessage("Wrong argument number.");
            return;
        }

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            try
            {
                TypedValue entity = rbArgs.AsArray().First();

                if (entity.TypeCode != (int)LispDataType.ObjectId)
                {
                    throw new System.Exception($"Wrong argument type: {entity.Value.ToString()}.");
                }

                ObjectId id = (ObjectId)entity.Value;
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);

                ObjectIdCollection blockRefIds = btr.GetBlockReferenceIds(true, true);

                foreach (ObjectId blockRefId in blockRefIds)
                {
                    BlockReference br = (BlockReference)tr.GetObject(blockRefId, OpenMode.ForRead);
                    SynchronizeAttributes(br, tr);
                }

                tr.Commit();
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(ex.Message);
                tr.Abort();
            }
        }
    }

    private void SynchronizeAttributes(BlockReference br, Transaction tr)
    {

        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
        if (btr.HasAttributeDefinitions)
        {
            List<AttributeDefinition> attDefs = GetAttributes(btr, tr);
            ResetAttributes(br, attDefs, tr);
        }
    }

    private List<AttributeDefinition> GetAttributes(BlockTableRecord btr, Transaction tr)
    {
        List<AttributeDefinition> attDefs = new List<AttributeDefinition>();
        foreach (ObjectId id in btr)
        {
            if (id.ObjectClass == attDefClass)
            {
                AttributeDefinition attDef = (AttributeDefinition)tr.GetObject(id, OpenMode.ForRead);
                attDefs.Add(attDef);
            }
        }

        return attDefs;
    }

    private void ResetAttributes(BlockReference br, List<AttributeDefinition> attDefs,Transaction tr)
    {
        Dictionary<string, string> attValues = new Dictionary<string, string>();
        foreach (ObjectId id in br.AttributeCollection)
        {
            AttributeReference attRef = (AttributeReference)tr.GetObject(id, OpenMode.ForWrite);
            string attValue = attRef.IsMTextAttribute ? attRef.MTextAttribute.Contents : attRef.TextString;
            attValues.Add(attRef.Tag, attValue);
            attRef.Erase();
        }

        br.UpgradeOpen();

        foreach (AttributeDefinition attDef in attDefs)
        {
            AttributeReference attRef = new AttributeReference();
            attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
            if (attDef.Constant)
            {
                string attValue = attDef.IsMTextAttributeDefinition ? attDef.MTextAttributeDefinition.Contents : attDef.TextString;
                attRef.TextString = attValue;
            }
            else if (attValues.ContainsKey(attDef.Tag))
            {
                attRef.TextString = attValues[attDef.Tag];
            }
            br.AttributeCollection.AppendAttribute(attRef);
            tr.AddNewlyCreatedDBObject(attRef, true);
        }
    }
}

