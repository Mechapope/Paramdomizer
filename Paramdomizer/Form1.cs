using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MeowDSIO;
using MeowDSIO.DataFiles;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace Paramdomizer
{
    public partial class Form1 : Form
    {
        string gameDirectory = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //check if running exe from data directory
            gameDirectory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

            if (File.Exists(gameDirectory + "\\DARKSOULS.exe"))
            {
                //exe is in a valid game directory, just use this as the path instead of asking for input
                txtGamePath.Text = gameDirectory;
                txtGamePath.ReadOnly = true;

                if (!File.Exists(gameDirectory + "\\param\\GameParam\\GameParam.parambnd"))
                {
                    //user hasn't unpacked their game
                    lblMessage.Text = "You don't seem to have an unpacked Dark Souls installation. Please run UDSFM and come back :)";
                    lblMessage.Visible = true;
                    lblMessage.ForeColor = Color.Red;
                }
            }
        }

        private void btnOpenFolderDialog_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                txtGamePath.Text = dialog.FileName;
                gameDirectory = dialog.FileName;

                lblMessage.Text = "";
                lblMessage.Visible = true;

                if (!File.Exists(gameDirectory + "\\DARKSOULS.exe"))
                {
                    lblMessage.Text = "Not a valid Data directory!";
                    lblMessage.ForeColor = Color.Red;
                    return;
                }
                else if (!File.Exists(gameDirectory + "\\param\\GameParam\\GameParam.parambnd"))
                {
                    //user hasn't unpacked their game
                    lblMessage.Text = "You don't seem to have an unpacked Dark Souls installation. Please run UDSFM and come back :)";
                    lblMessage.ForeColor = Color.Red;
                    return;
                }
            }
        }

        private async void btnSubmit_Click(object sender, EventArgs e)
        {
            //check that entered path is valid
            gameDirectory = txtGamePath.Text;

            //reset message label
            lblMessage.Text = "";
            lblMessage.ForeColor = new Color();
            lblMessage.Visible = true;

            if (!File.Exists(gameDirectory + "\\DARKSOULS.exe"))
            {
                lblMessage.Text = "Not a valid Data directory!";
                lblMessage.ForeColor = Color.Red;
                return;
            }
            else if (!File.Exists(gameDirectory + "\\param\\GameParam\\GameParam.parambnd"))
            {
                //user hasn't unpacked their game
                lblMessage.Text = "You don't seem to have an unpacked Dark Souls installation. Please run UDSFM and come back :)";
                lblMessage.ForeColor = Color.Red;
                return;
            }

            //update label on a new thread
            Progress<string> progress = new Progress<string>(s => lblMessage.Text = s);
            await Task.Factory.StartNew(() => UiThread.WriteToInfoLabel(progress));

            //generate a seed if needed
            if (txtSeed.Text == "")
            {
                string validChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
                Random seedGen = new Random();
                for (int i = 0; i < 15; i++)
                {
                    txtSeed.Text += validChars[seedGen.Next(validChars.Length)];
                }
            }

            string seed = txtSeed.Text;

            //create backup of gameparam
            if (!File.Exists(gameDirectory + "\\param\\GameParam\\GameParam.parambnd.bak"))
            {
                File.Copy(gameDirectory + "\\param\\GameParam\\GameParam.parambnd", gameDirectory + "\\param\\GameParam\\GameParam.parambnd.bak");
                lblMessage.Text = "Backed up GameParam.parambnd \n\n";
                lblMessage.ForeColor = Color.Black;
                lblMessage.Visible = true;
            }

            //Load parambnds/paramdefs into params
            List<PARAM> AllParams = new List<PARAM>();
            List<PARAMDEF> ParamDefs = new List<PARAMDEF>();

            List<BND> gameparamBnds = Directory.GetFiles(gameDirectory + "\\param\\GameParam\\", "*.parambnd")
                .Select(p => DataFile.LoadFromFile<BND>(p, new Progress<(int, int)>((pr) =>
                {

                }))).ToList();

            List<BND> paramdefBnds = Directory.GetFiles(gameDirectory + "\\paramdef\\", "*.paramdefbnd")
                .Select(p => DataFile.LoadFromFile<BND>(p, new Progress<(int, int)>((pr) =>
                {

                }))).ToList();

            for (int i = 0; i < paramdefBnds.Count(); i++)
            {
                foreach (MeowDSIO.DataTypes.BND.BNDEntry paramdef in paramdefBnds[i])
                {
                    PARAMDEF newParamDef = paramdef.ReadDataAs<PARAMDEF>(new Progress<(int, int)>((p) =>
                    {

                    }));
                    ParamDefs.Add(newParamDef);
                }
            }

            for (int i = 0; i < gameparamBnds.Count(); i++)
            {
                foreach (MeowDSIO.DataTypes.BND.BNDEntry param in gameparamBnds[i])
                {
                    PARAM newParam = param.ReadDataAs<PARAM>(new Progress<(int, int)>((p) =>
                    {

                    }));

                    newParam.ApplyPARAMDEFTemplate(ParamDefs.Where(x => x.ID == newParam.ID).First());
                    AllParams.Add(newParam);
                }
            }

            //Hash seed so people can use meme seeds
            Random r = new Random(seed.GetHashCode());

            foreach (PARAM paramFile in AllParams)
            {
                if (paramFile.VirtualUri.EndsWith("AtkParam_Npc.param"))
                {
                    List<int> allSpEffects = new List<int>();
                    List<float> allKnockbackDists = new List<float>();
                    List<int> allDmgLevels = new List<int>();
                    List<float> allHitRadius = new List<float>();

                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "spEffectId0")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allSpEffects.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "knockbackDist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allKnockbackDists.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "dmgLevel")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allDmgLevels.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "hit0_Radius")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allHitRadius.Add((float)(prop.GetValue(cell, null)));
                            }
                        }
                    }

                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "knockbackDist")
                            {
                                int randomIndex = r.Next(allKnockbackDists.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkKnockback.Checked)
                                {
                                    prop.SetValue(cell, allKnockbackDists[randomIndex], null);
                                }

                                allKnockbackDists.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "dmgLevel")
                            {
                                int randomIndex = r.Next(allDmgLevels.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkStaggerLevels.Checked)
                                {
                                    prop.SetValue(cell, allDmgLevels[randomIndex], null);
                                }

                                allDmgLevels.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "hit0_Radius")
                            {
                                int randomIndex = r.Next(allHitRadius.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkHitboxSizes.Checked)
                                {
                                    prop.SetValue(cell, allHitRadius[randomIndex], null);
                                }

                                allHitRadius.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "spEffectId0")
                            {
                                int randomIndex = r.Next(allSpEffects.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkSpeffects.Checked)
                                {
                                    prop.SetValue(cell, allSpEffects[randomIndex], null);
                                }

                                allSpEffects.RemoveAt(randomIndex);
                            }
                        }
                    }
                }
                else if (paramFile.VirtualUri.EndsWith("AtkParam_Pc.param"))
                {
                    List<float> allKnockbackDists = new List<float>();
                    List<int> allDmgLevels = new List<int>();
                    List<float> allHitRadius = new List<float>();

                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "knockbackDist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allKnockbackDists.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "dmgLevel")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allDmgLevels.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "hit0_Radius")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allHitRadius.Add((float)(prop.GetValue(cell, null)));
                            }
                        }
                    }

                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "knockbackDist")
                            {
                                int randomIndex = r.Next(allKnockbackDists.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkKnockback.Checked)
                                {
                                    prop.SetValue(cell, allKnockbackDists[randomIndex], null);
                                }

                                allKnockbackDists.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "dmgLevel")
                            {
                                int randomIndex = r.Next(allDmgLevels.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkStaggerLevels.Checked)
                                {
                                    prop.SetValue(cell, allDmgLevels[randomIndex], null);
                                }

                                allDmgLevels.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "hit0_Radius")
                            {
                                int randomIndex = r.Next(allHitRadius.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkHitboxSizes.Checked)
                                {
                                    prop.SetValue(cell, allHitRadius[randomIndex], null);
                                }

                                allHitRadius.RemoveAt(randomIndex);
                            }
                        }
                    }
                }
                else if (paramFile.ID == "NPC_PARAM_ST")
                {
                    List<int> allSPeffects = new List<int>();
                    List<float> allTurnVelocities = new List<float>();
                    List<int> allStaminas = new List<int>();
                    List<int> allStaminaRegens = new List<int>();

                    //Dont randomize these speffects
                    int[] invalidSpeffects = { 0, 5300, 7001, 7002, 7003, 7004, 7005, 7006, 7007, 7008, 7009, 7010, 7011, 7012, 7013, 7014, 7015 };

                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            //spEffect rando
                            if (cell.Def.Name.StartsWith("spEffectID"))
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                int speffectCheck = Convert.ToInt32(prop.GetValue(cell, null));
                                if (!invalidSpeffects.Contains(speffectCheck))
                                {
                                    allSPeffects.Add(speffectCheck);
                                }
                            }
                            else if (cell.Def.Name == "turnVellocity")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allTurnVelocities.Add((float)(prop.GetValue(cell, null)));
                            }
                            /*else if (cell.Def.Name == "stamina")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allStaminas.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "staminaRecoverBaseVel")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allStaminaRegens.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }*/
                        }
                    }

                    //loop again to set a random value per entry
                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name.StartsWith("spEffectID"))
                            {
                                int randomIndex = r.Next(allSPeffects.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                int speffectCheck = Convert.ToInt32(prop.GetValue(cell, null));

                                if (!invalidSpeffects.Contains(speffectCheck))
                                {
                                    if (chkSpeffects.Checked)
                                    {
                                        prop.SetValue(cell, allSPeffects[randomIndex], null);
                                    }
                                    allSPeffects.RemoveAt(randomIndex);
                                }
                            }
                            else if (cell.Def.Name == "turnVellocity")
                            {
                                int randomIndex = r.Next(allTurnVelocities.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkTurnSpeeds.Checked)
                                {
                                    prop.SetValue(cell, allTurnVelocities[randomIndex], null);
                                }

                                allTurnVelocities.RemoveAt(randomIndex);
                            }                             
                            //else if (cell.Def.Name == "stamina")
                            //{
                            //    int randomIndex = r.Next(allStaminas.Count);
                            //    Type type = cell.GetType();
                            //    PropertyInfo prop = type.GetProperty("Value");

                            //    if (chkStaminaRegen.Checked)
                            //    {
                            //        prop.SetValue(cell, allStaminas[randomIndex], null);
                            //    }

                            //    allStaminas.RemoveAt(randomIndex);
                            //}
                            //else if (cell.Def.Name == "staminaRecoverBaseVal")
                            //{
                            //    int randomIndex = r.Next(allStaminaRegens.Count);
                            //    Type type = cell.GetType();
                            //    PropertyInfo prop = type.GetProperty("Value");

                            //    if (chkStaminaRegen.Checked)
                            //    {
                            //        prop.SetValue(cell, allStaminaRegens[randomIndex], null);
                            //    }

                            //    allStaminaRegens.RemoveAt(randomIndex);
                            //}
                        }
                    }
                }
                else if (paramFile.ID == "NPC_THINK_PARAM_ST")
                {
                    List<float> allNearDists = new List<float>();
                    List<float> allMidDists = new List<float>();
                    List<float> allFarDists = new List<float>();
                    List<float> allOutDists = new List<float>();
                    List<int> allEye_dists = new List<int>();
                    List<int> allEar_dists = new List<int>();
                    List<int> allNose_dists = new List<int>();
                    List<int> allMaxBackhomeDists = new List<int>();
                    List<int> allBackhomeDists = new List<int>();
                    List<int> allBackhomeBattleDists = new List<int>();
                    List<int> allBackHome_LookTargetTimes = new List<int>();
                    List<int> allBackHome_LookTargetDists = new List<int>();
                    List<int> allBattleStartDists = new List<int>();
                    List<int> allEye_angXs = new List<int>();
                    List<int> allEye_angYs = new List<int>();
                    List<int> allEar_angXs = new List<int>();
                    List<int> allEar_angYs = new List<int>();
                    List<int> allSightTargetForgetTimes = new List<int>();
                    List<int> allSoundTargetForgetTimes = new List<int>();

                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "nearDist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allNearDists.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "midDist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allMidDists.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "farDist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allFarDists.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "outDist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allOutDists.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "eye_dist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allEye_dists.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "ear_dist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allEar_dists.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "nose_dist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allNose_dists.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "maxBackhomeDist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allMaxBackhomeDists.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "backhomeDist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allBackhomeDists.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "backhomeBattleDist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allBackhomeBattleDists.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "BackHome_LookTargetTime")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allBackHome_LookTargetTimes.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "BackHome_LookTargetDist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allBackHome_LookTargetDists.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "BattleStartDist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allBattleStartDists.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "eye_angX")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allEye_angXs.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "eye_angY")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allEye_angYs.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "ear_angX")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allEar_angXs.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "ear_angY")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allEar_angYs.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "SightTargetForgetTime")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allSightTargetForgetTimes.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "SoundTargetForgetTime")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allSoundTargetForgetTimes.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                        }
                    }
                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            int[] hydraIds = { 353000, 353001, 353100, 353200 };
                            if (cell.Def.Name == "nearDist")
                            {                                
                                int randomIndex = r.Next(allNearDists.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allNearDists[randomIndex], null);
                                }

                                allNearDists.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "midDist")
                            {
                                int randomIndex = r.Next(allMidDists.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allMidDists[randomIndex], null);
                                }

                                allMidDists.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "farDist")
                            {
                                int randomIndex = r.Next(allFarDists.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allFarDists[randomIndex], null);
                                }

                                allFarDists.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "outDist")
                            {
                                int randomIndex = r.Next(allOutDists.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allOutDists[randomIndex], null);
                                }

                                allOutDists.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "eye_dist")
                            {
                                int randomIndex = r.Next(allEye_dists.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allEye_dists[randomIndex], null);
                                }

                                allEye_dists.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "ear_dist")
                            {
                                int randomIndex = r.Next(allEar_dists.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allEar_dists[randomIndex], null);
                                }

                                allEar_dists.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "nose_dist")
                            {
                                int randomIndex = r.Next(allNose_dists.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allNose_dists[randomIndex], null);
                                }

                                allNose_dists.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "maxBackhomeDist")
                            {
                                int randomIndex = r.Next(allMaxBackhomeDists.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allMaxBackhomeDists[randomIndex], null);
                                }

                                allMaxBackhomeDists.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "backhomeDist")
                            {
                                int randomIndex = r.Next(allBackhomeDists.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allBackhomeDists[randomIndex], null);
                                }

                                allBackhomeDists.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "backhomeBattleDist")
                            {
                                int randomIndex = r.Next(allBackhomeBattleDists.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allBackhomeBattleDists[randomIndex], null);
                                }

                                allBackhomeBattleDists.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "BackHome_LookTargetTime")
                            {
                                int randomIndex = r.Next(allBackHome_LookTargetTimes.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allBackHome_LookTargetTimes[randomIndex], null);
                                }

                                allBackHome_LookTargetTimes.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "BackHome_LookTargetDist")
                            {
                                int randomIndex = r.Next(allBackHome_LookTargetDists.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allBackHome_LookTargetDists[randomIndex], null);
                                }

                                allBackHome_LookTargetDists.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "BattleStartDist")
                            {
                                int randomIndex = r.Next(allBattleStartDists.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allBattleStartDists[randomIndex], null);
                                }

                                allBattleStartDists.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "eye_angX")
                            {
                                int randomIndex = r.Next(allEye_angXs.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allEye_angXs[randomIndex], null);
                                }

                                allEye_angXs.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "eye_angY")
                            {
                                int randomIndex = r.Next(allEye_angYs.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allEye_angYs[randomIndex], null);
                                }

                                allEye_angYs.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "ear_angX")
                            {
                                int randomIndex = r.Next(allEar_angXs.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allEar_angXs[randomIndex], null);
                                }

                                allEar_angXs.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "ear_angY")
                            {
                                int randomIndex = r.Next(allEar_angYs.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allEar_angYs[randomIndex], null);
                                }

                                allEar_angYs.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "SightTargetForgetTime")
                            {
                                int randomIndex = r.Next(allSightTargetForgetTimes.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allSightTargetForgetTimes[randomIndex], null);
                                }

                                allSightTargetForgetTimes.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "SoundTargetForgetTime")
                            {
                                int randomIndex = r.Next(allSoundTargetForgetTimes.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked && !hydraIds.Contains(paramRow.ID))
                                {
                                    prop.SetValue(cell, allSoundTargetForgetTimes[randomIndex], null);
                                }

                                allSoundTargetForgetTimes.RemoveAt(randomIndex);
                            }
                        }
                    }
                }
                else if (paramFile.ID == "EQUIP_PARAM_ACCESSORY_ST")
                {
                    List<int> allRefIds = new List<int>();
                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "refId")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allRefIds.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                        }
                    }

                    //loop again to set a random value per entry
                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "refId")
                            {
                                int randomIndex = r.Next(allRefIds.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkRingSpeffects.Checked)
                                {
                                    prop.SetValue(cell, allRefIds[randomIndex], null);
                                }

                                allRefIds.RemoveAt(randomIndex);
                            }
                        }
                    }
                }
                else if (paramFile.ID == "EQUIP_PARAM_GOODS_ST")
                {
                    List<int> allUseAnimations = new List<int>();
                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "goodsUseAnim")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                int animid = Convert.ToInt32(prop.GetValue(cell, null));
                                //the empty estus animation - prevents using item
                                if (animid != 254)
                                {
                                    allUseAnimations.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                                }
                            }
                        }
                    }

                    //loop again to set a random value per entry
                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "goodsUseAnim")
                            {
                                int randomIndex = r.Next(allUseAnimations.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");
                                int animid = Convert.ToInt32(prop.GetValue(cell, null));

                                if (animid != 254)
                                {
                                    if (chkItemAnimations.Checked)
                                    {
                                        prop.SetValue(cell, allUseAnimations[randomIndex], null);
                                    }

                                    allUseAnimations.RemoveAt(randomIndex);
                                }
                            }
                        }
                    }
                }
                else if (paramFile.ID == "EQUIP_PARAM_WEAPON_ST")
                {
                    //loop through all entries once to get list of values
                    List<int> allWepmotionCats = new List<int>();
                    List<int> allWepmotion1hCats = new List<int>();
                    List<int> allWepmotion2hCats = new List<int>();
                    List<int> allspAtkcategories = new List<int>();
                    List<int> allAttackBasePhysic = new List<int>();
                    List<int> allAttackBaseMagic = new List<int>();
                    List<int> allAttackBaseFire = new List<int>();
                    List<int> allAttackBaseThunder = new List<int>();
                    List<int> allEquipModelIds = new List<int>();

                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        //check not to randomize moveset of bows which defeats the purpose of bullet rando. Try disabling this check and see what happens maybe
                        MeowDSIO.DataTypes.PARAM.ParamCellValueRef bowCheckCell = paramRow.Cells.First(c => c.Def.Name == "bowDistRate");
                        Type bowchecktype = bowCheckCell.GetType();
                        PropertyInfo bowcheckprop = bowchecktype.GetProperty("Value");
                        if (Convert.ToInt32(bowcheckprop.GetValue(bowCheckCell, null)) < 0)
                        {
                            foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                            {
                                if (cell.Def.Name == "wepmotionCategory")
                                {
                                    PropertyInfo prop = cell.GetType().GetProperty("Value");
                                    allWepmotionCats.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                                }
                                else if (cell.Def.Name == "wepmotionOneHandId")
                                {
                                    PropertyInfo prop = cell.GetType().GetProperty("Value");
                                    allWepmotion1hCats.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                                }
                                else if (cell.Def.Name == "wepmotionBothHandId")
                                {
                                    PropertyInfo prop = cell.GetType().GetProperty("Value");
                                    allWepmotion2hCats.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                                }
                                else if (cell.Def.Name == "spAtkcategory")
                                {
                                    PropertyInfo prop = cell.GetType().GetProperty("Value");
                                    allspAtkcategories.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                                }
                            }
                        }

                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "equipModelId")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allEquipModelIds.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "attackBasePhysics")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allAttackBasePhysic.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "attackBaseMagic")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allAttackBaseMagic.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "attackBaseFire")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allAttackBaseFire.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "attackBaseThunder")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allAttackBaseThunder.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                        }
                    }

                    //loop again to set a random value per entry
                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        MeowDSIO.DataTypes.PARAM.ParamCellValueRef bowCheckCell = paramRow.Cells.First(c => c.Def.Name == "bowDistRate");
                        Type bowchecktype = bowCheckCell.GetType();
                        PropertyInfo bowcheckprop = bowchecktype.GetProperty("Value");

                        if (Convert.ToInt32(bowcheckprop.GetValue(bowCheckCell, null)) < 0)
                        {
                            foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                            {
                                if (cell.Def.Name == "wepmotionCategory")
                                {
                                    int randomIndex = r.Next(allWepmotionCats.Count);
                                    Type type = cell.GetType();
                                    PropertyInfo prop = type.GetProperty("Value");
                                    if (chkWeaponMoveset.Checked)
                                    {
                                        prop.SetValue(cell, allWepmotionCats[randomIndex], null);
                                    }

                                    allWepmotionCats.RemoveAt(randomIndex);
                                }
                                else if (cell.Def.Name == "wepmotionOneHandId")
                                {
                                    int randomIndex = r.Next(allWepmotion1hCats.Count);
                                    Type type = cell.GetType();
                                    PropertyInfo prop = type.GetProperty("Value");
                                    if (chkWeaponMoveset.Checked)
                                    {
                                        prop.SetValue(cell, allWepmotion1hCats[randomIndex], null);
                                    }

                                    allWepmotion1hCats.RemoveAt(randomIndex);
                                }
                                else if (cell.Def.Name == "wepmotionBothHandId")
                                {
                                    int randomIndex = r.Next(allWepmotion2hCats.Count);
                                    Type type = cell.GetType();
                                    PropertyInfo prop = type.GetProperty("Value");
                                    if (chkWeaponMoveset.Checked)
                                    {
                                        prop.SetValue(cell, allWepmotion2hCats[randomIndex], null);
                                    }

                                    allWepmotion2hCats.RemoveAt(randomIndex);
                                }
                                else if (cell.Def.Name == "spAtkcategory")
                                {
                                    int randomIndex = r.Next(allspAtkcategories.Count);
                                    Type type = cell.GetType();
                                    PropertyInfo prop = type.GetProperty("Value");

                                    if (chkWeaponMoveset.Checked)
                                    {
                                        prop.SetValue(cell, allspAtkcategories[randomIndex], null);
                                    }

                                    allspAtkcategories.RemoveAt(randomIndex);
                                }
                            }
                        }

                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "equipModelId")
                            {
                                int randomIndex = r.Next(allEquipModelIds.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");
                                if (chkWeaponModels.Checked)
                                {
                                    prop.SetValue(cell, allEquipModelIds[randomIndex], null);
                                }

                                allEquipModelIds.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "attackBasePhysics")
                            {
                                int randomIndex = r.Next(allAttackBasePhysic.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");
                                if (chkWeaponDamage.Checked)
                                {
                                    prop.SetValue(cell, allAttackBasePhysic[randomIndex], null);
                                }

                                allAttackBasePhysic.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "attackBaseMagic")
                            {
                                int randomIndex = r.Next(allAttackBaseMagic.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");
                                if (chkWeaponDamage.Checked)
                                {
                                    prop.SetValue(cell, allAttackBaseMagic[randomIndex], null);
                                }

                                allAttackBaseMagic.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "attackBaseFire")
                            {
                                int randomIndex = r.Next(allAttackBaseFire.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");
                                if (chkWeaponDamage.Checked)
                                {
                                    prop.SetValue(cell, allAttackBaseFire[randomIndex], null);
                                }

                                allAttackBaseFire.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "attackBaseThunder")
                            {
                                int randomIndex = r.Next(allAttackBaseThunder.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");
                                if (chkWeaponDamage.Checked)
                                {
                                    prop.SetValue(cell, allAttackBaseThunder[randomIndex], null);
                                }

                                allAttackBaseThunder.RemoveAt(randomIndex);
                            }
                        }
                    }
                }
                else if (paramFile.ID == "MAGIC_PARAM_ST")
                {
                    //loop through all entries once to get list of all values
                    List<int> allSfxVariationIds = new List<int>();
                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "refType")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allSfxVariationIds.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                        }
                    }

                    //loop again to set a random value per entry
                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "refType")
                            {
                                int randomIndex = r.Next(allSfxVariationIds.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkMagicAnimations.Checked)
                                {
                                    prop.SetValue(cell, allSfxVariationIds[randomIndex], null);
                                }

                                allSfxVariationIds.RemoveAt(randomIndex);
                            }
                        }
                    }
                }
                else if (paramFile.ID == "TALK_PARAM_ST")
                {
                    //loop through all entries once to get list of all values
                    List<int> allSounds = new List<int>();
                    List<int> allMsgs = new List<int>();
                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "voiceId")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                if (!InvalidVoiceIds.Contains(prop.GetValue(cell, null)))
                                {
                                    allSounds.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                                }
                            }
                            else if (cell.Def.Name == "msgId")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                if (!InvalidVoiceIds.Contains(prop.GetValue(cell, null)))
                                {
                                    allMsgs.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                                }
                            }
                        }
                    }

                    //loop again to set a random value per entry
                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "voiceId")
                            {
                                int randomIndex = r.Next(allSounds.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkVoices.Checked && !InvalidVoiceIds.Contains(prop.GetValue(cell, null)))
                                {
                                    prop.SetValue(cell, allSounds[randomIndex], null);
                                }

                                allSounds.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "msgId")
                            {
                                int randomIndex = r.Next(allMsgs.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");
                                if (chkVoices.Checked && !InvalidVoiceIds.Contains(prop.GetValue(cell, null)))
                                {
                                    prop.SetValue(cell, allMsgs[randomIndex], null);
                                }

                                allMsgs.RemoveAt(randomIndex);
                            }
                        }
                    }
                }
                else if (paramFile.ID == "RAGDOLL_PARAM_ST")
                {
                    //dunno if this one actually does anything
                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            //TODO...maybe
                        }
                    }
                }
                else if (paramFile.ID == "SKELETON_PARAM_ST")
                {
                    List<int> allneckTurnGains = new List<int>();
                    List<int> alloriginalGroundHeightMSs = new List<int>();
                    List<int> allminAnkleHeightMSs = new List<int>();
                    List<int> allmaxAnkleHeightMSs = new List<int>();
                    List<int> allcosineMaxKneeAngles = new List<int>();
                    List<int> allcosineMinKneeAngles = new List<int>();
                    List<int> allfootPlantedAnkleHeightMSs = new List<int>();
                    List<int> allfootRaisedAnkleHeightMSs = new List<int>();
                    List<int> allrayCastDistanceUps = new List<int>();
                    List<int> allraycastDistanceDowns = new List<int>();
                    List<int> allfootEndLS_Xs = new List<int>();
                    List<int> allfootEndLS_Ys = new List<int>();
                    List<int> allfootEndLS_Zs = new List<int>();
                    List<int> allonOffGains = new List<int>();
                    List<int> allgroundAscendingGains = new List<int>();
                    List<int> allgroundDescendingGains = new List<int>();
                    List<int> allfootRaisedGains = new List<int>();
                    List<int> allfootDescendingGains = new List<int>();
                    List<int> allfootUnlockGains = new List<int>();
                    List<int> allkneeAxisTypes = new List<int>();
                    List<int> alluseFootLockings = new List<int>();
                    List<int> allfootPlacementOns = new List<int>();
                    List<int> alltwistKneeAxisTypes = new List<int>();
                    List<int> allneckTurnPrioritys = new List<int>();
                    List<int> allneckTurnMaxAngles = new List<int>();

                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "neckTurnGain")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allneckTurnGains.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "originalGroundHeightMS")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                alloriginalGroundHeightMSs.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "minAnkleHeightMS")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allminAnkleHeightMSs.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "maxAnkleHeightMS")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allmaxAnkleHeightMSs.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "cosineMaxKneeAngle")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allcosineMaxKneeAngles.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "cosineMinKneeAngle")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allcosineMinKneeAngles.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "footPlantedAnkleHeightMS")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allfootPlantedAnkleHeightMSs.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "footRaisedAnkleHeightMS")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allfootRaisedAnkleHeightMSs.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "rayCastDistanceUp")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allrayCastDistanceUps.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "raycastDistanceDown")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allraycastDistanceDowns.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "footEndLS_X")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allfootEndLS_Xs.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "footEndLS_Y")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allfootEndLS_Ys.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "footEndLS_Z")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allfootEndLS_Zs.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "onOffGain")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allonOffGains.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "groundAscendingGain")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allgroundAscendingGains.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "groundDescendingGain")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allgroundDescendingGains.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "footRaisedGain")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allfootRaisedGains.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "footDescendingGain")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allfootDescendingGains.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "footUnlockGain")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allfootUnlockGains.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "kneeAxisType")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allkneeAxisTypes.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "useFootLocking")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                alluseFootLockings.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "footPlacementOn")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allfootPlacementOns.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "twistKneeAxisType")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                alltwistKneeAxisTypes.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "neckTurnPriority")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allneckTurnPrioritys.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                            else if (cell.Def.Name == "neckTurnMaxAngle")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allneckTurnMaxAngles.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }

                        }
                    }

                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "neckTurnGain")
                            {
                                int randomIndex = r.Next(allneckTurnGains.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if(true)
                                {
                                    prop.SetValue(cell, allneckTurnGains[randomIndex], null);
                                }

                                allneckTurnGains.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "originalGroundHeightMS")
                            {
                                int randomIndex = r.Next(alloriginalGroundHeightMSs.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, alloriginalGroundHeightMSs[randomIndex], null);
                                }

                                alloriginalGroundHeightMSs.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "minAnkleHeightMS")
                            {
                                int randomIndex = r.Next(allminAnkleHeightMSs.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allminAnkleHeightMSs[randomIndex], null);
                                }

                                allminAnkleHeightMSs.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "maxAnkleHeightMS")
                            {
                                int randomIndex = r.Next(allmaxAnkleHeightMSs.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allmaxAnkleHeightMSs[randomIndex], null);
                                }

                                allmaxAnkleHeightMSs.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "cosineMaxKneeAngle")
                            {
                                int randomIndex = r.Next(allcosineMaxKneeAngles.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allcosineMaxKneeAngles[randomIndex], null);
                                }

                                allcosineMaxKneeAngles.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "cosineMinKneeAngle")
                            {
                                int randomIndex = r.Next(allcosineMinKneeAngles.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allcosineMinKneeAngles[randomIndex], null);
                                }

                                allcosineMinKneeAngles.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "footPlantedAnkleHeightMS")
                            {
                                int randomIndex = r.Next(allfootPlantedAnkleHeightMSs.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allfootPlantedAnkleHeightMSs[randomIndex], null);
                                }

                                allfootPlantedAnkleHeightMSs.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "footRaisedAnkleHeightMS")
                            {
                                int randomIndex = r.Next(allfootRaisedAnkleHeightMSs.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allfootRaisedAnkleHeightMSs[randomIndex], null);
                                }

                                allfootRaisedAnkleHeightMSs.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "rayCastDistanceUp")
                            {
                                int randomIndex = r.Next(allrayCastDistanceUps.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allrayCastDistanceUps[randomIndex], null);
                                }

                                allrayCastDistanceUps.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "raycastDistanceDown")
                            {
                                int randomIndex = r.Next(allraycastDistanceDowns.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allraycastDistanceDowns[randomIndex], null);
                                }

                                allraycastDistanceDowns.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "footEndLS_X")
                            {
                                int randomIndex = r.Next(allfootEndLS_Xs.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allfootEndLS_Xs[randomIndex], null);
                                }

                                allfootEndLS_Xs.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "footEndLS_Y")
                            {
                                int randomIndex = r.Next(allfootEndLS_Ys.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allfootEndLS_Ys[randomIndex], null);
                                }

                                allfootEndLS_Ys.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "footEndLS_Z")
                            {
                                int randomIndex = r.Next(allfootEndLS_Zs.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allfootEndLS_Zs[randomIndex], null);
                                }

                                allfootEndLS_Zs.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "onOffGain")
                            {
                                int randomIndex = r.Next(allonOffGains.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allonOffGains[randomIndex], null);
                                }

                                allonOffGains.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "groundAscendingGain")
                            {
                                int randomIndex = r.Next(allgroundAscendingGains.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allgroundAscendingGains[randomIndex], null);
                                }

                                allgroundAscendingGains.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "groundDescendingGain")
                            {
                                int randomIndex = r.Next(allgroundDescendingGains.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allgroundDescendingGains[randomIndex], null);
                                }

                                allgroundDescendingGains.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "footRaisedGain")
                            {
                                int randomIndex = r.Next(allfootRaisedGains.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allfootRaisedGains[randomIndex], null);
                                }

                                allfootRaisedGains.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "footDescendingGain")
                            {
                                int randomIndex = r.Next(allfootDescendingGains.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allfootDescendingGains[randomIndex], null);
                                }

                                allfootDescendingGains.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "footUnlockGain")
                            {
                                int randomIndex = r.Next(allfootUnlockGains.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allfootUnlockGains[randomIndex], null);
                                }

                                allfootUnlockGains.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "kneeAxisType")
                            {
                                int randomIndex = r.Next(allkneeAxisTypes.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allkneeAxisTypes[randomIndex], null);
                                }

                                allkneeAxisTypes.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "useFootLocking")
                            {
                                int randomIndex = r.Next(alluseFootLockings.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, alluseFootLockings[randomIndex], null);
                                }

                                alluseFootLockings.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "footPlacementOn")
                            {
                                int randomIndex = r.Next(allfootPlacementOns.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allfootPlacementOns[randomIndex], null);
                                }

                                allfootPlacementOns.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "twistKneeAxisType")
                            {
                                int randomIndex = r.Next(alltwistKneeAxisTypes.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, alltwistKneeAxisTypes[randomIndex], null);
                                }

                                alltwistKneeAxisTypes.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "neckTurnPriority")
                            {
                                int randomIndex = r.Next(allneckTurnPrioritys.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allneckTurnPrioritys[randomIndex], null);
                                }

                                allneckTurnPrioritys.RemoveAt(randomIndex);
                            }

                            else if (cell.Def.Name == "neckTurnMaxAngle")
                            {
                                int randomIndex = r.Next(allneckTurnMaxAngles.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                //if (chkSkeletons.Checked)
                                if (true)
                                {
                                    prop.SetValue(cell, allneckTurnMaxAngles[randomIndex], null);
                                }

                                allneckTurnMaxAngles.RemoveAt(randomIndex);
                            }

                        }
                    }
                }
                else if (paramFile.ID == "BULLET_PARAM_ST")
                {
                    //build a list of all properties to rando
                    List<int> atkId_BulletVals = new List<int>();
                    List<int> sfxId_BulletVals = new List<int>();
                    List<int> sfxId_HitVals = new List<int>();
                    List<int> sfxId_FlickVals = new List<int>();
                    List<float> lifeVals = new List<float>();
                    List<float> distVals = new List<float>();
                    List<float> shootIntervalVals = new List<float>();
                    List<float> gravityInRangeVals = new List<float>();
                    List<float> gravityOutRangeVals = new List<float>();
                    List<float> hormingStopRangeVals = new List<float>();
                    List<float> initVellocityVals = new List<float>();
                    List<float> accelInRangeVals = new List<float>();
                    List<float> accelOutRangeVals = new List<float>();
                    List<float> maxVellocityVals = new List<float>();
                    List<float> minVellocityVals = new List<float>();
                    List<float> accelTimeVals = new List<float>();
                    List<float> homingBeginDistVals = new List<float>();
                    List<float> hitRadiusVals = new List<float>();
                    List<float> hitRadiusMaxVals = new List<float>();
                    List<float> spreadTimeVals = new List<float>();
                    List<float> hormingOffsetRangeVals = new List<float>();
                    List<float> dmgHitRecordLifeTimeVals = new List<float>();
                    List<int> spEffectIDForShooterVals = new List<int>();
                    List<int> HitBulletIDVals = new List<int>();
                    List<int> spEffectId0Vals = new List<int>();
                    List<ushort> numShootVals = new List<ushort>();
                    List<short> homingAngleVals = new List<short>();
                    List<short> shootAngleVals = new List<short>();
                    List<short> shootAngleIntervalVals = new List<short>();
                    List<short> shootAngleXIntervalVals = new List<short>();
                    List<sbyte> damageDampVals = new List<sbyte>();
                    List<sbyte> spelDamageDampVals = new List<sbyte>();
                    List<sbyte> fireDamageDampVals = new List<sbyte>();
                    List<sbyte> thunderDamageDampVals = new List<sbyte>();
                    List<sbyte> staminaDampVals = new List<sbyte>();
                    List<sbyte> knockbackDampVals = new List<sbyte>();
                    List<sbyte> shootAngleXZVals = new List<sbyte>();
                    List<int> lockShootLimitAngVals = new List<int>();
                    List<int> isPenetrateVals = new List<int>();
                    List<int> atkAttributeVals = new List<int>();
                    List<int> spAttributeVals = new List<int>();
                    List<int> Material_AttackTypeVals = new List<int>();
                    List<int> Material_AttackMaterialVals = new List<int>();
                    List<int> Material_SizeVals = new List<int>();
                    List<int> launchConditionTypeVals = new List<int>();
                    List<int> FollowTypeVals = new List<int>();
                    List<int> isAttackSFXVals = new List<int>();
                    List<int> isEndlessHitVals = new List<int>();
                    List<int> isPenetrateMapVals = new List<int>();
                    List<int> isHitBothTeamVals = new List<int>();
                    List<int> isUseSharedHitListVals = new List<int>();
                    List<int> isHitForceMagicVals = new List<int>();
                    List<int> isIgnoreSfxIfHitWaterVals = new List<int>();
                    List<int> IsIgnoreMoveStateIfHitWaterVals = new List<int>();
                    List<int> isHitDarkForceMagicVals = new List<int>();

                    //add to list to rando
                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "atkId_Bullet")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                if (Convert.ToInt32(prop.GetValue(cell, null)) > 0)
                                {
                                    atkId_BulletVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                                }
                            }
                            else if (cell.Def.Name == "sfxId_Bullet")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                sfxId_BulletVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "sfxId_Hit")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                sfxId_HitVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "sfxId_Flick")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                sfxId_FlickVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "life")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                lifeVals.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "dist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                if ((float)(prop.GetValue(cell, null)) > 0)
                                {
                                    distVals.Add((float)(prop.GetValue(cell, null)));
                                }
                            }
                            else if (cell.Def.Name == "shootInterval")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                if ((float)(prop.GetValue(cell, null)) > 0)
                                {
                                    shootIntervalVals.Add((float)(prop.GetValue(cell, null)));
                                }
                            }
                            else if (cell.Def.Name == "gravityInRange")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                gravityInRangeVals.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "gravityOutRange")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                gravityOutRangeVals.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "hormingStopRange")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                hormingStopRangeVals.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "initVellocity")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                if ((float)(prop.GetValue(cell, null)) > 0)
                                {
                                    initVellocityVals.Add((float)(prop.GetValue(cell, null)));
                                }
                            }
                            else if (cell.Def.Name == "accelInRange")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                accelInRangeVals.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "accelOutRange")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                accelOutRangeVals.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "maxVellocity")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                if ((float)(prop.GetValue(cell, null)) > 0)
                                {
                                    maxVellocityVals.Add((float)(prop.GetValue(cell, null)));
                                }
                            }
                            else if (cell.Def.Name == "minVellocity")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                minVellocityVals.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "accelTime")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                accelTimeVals.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "homingBeginDist")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                homingBeginDistVals.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "hitRadius")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                hitRadiusVals.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "hitRadiusMax")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                hitRadiusMaxVals.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "spreadTime")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                spreadTimeVals.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "hormingOffsetRange")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                hormingOffsetRangeVals.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "dmgHitRecordLifeTime")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                dmgHitRecordLifeTimeVals.Add((float)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "spEffectIDForShooter")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                spEffectIDForShooterVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "HitBulletID")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                HitBulletIDVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "spEffectId0")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                spEffectId0Vals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "numShoot")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                numShootVals.Add((ushort)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "homingAngle")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                homingAngleVals.Add((short)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "shootAngle")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                shootAngleVals.Add((short)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "shootAngleInterval")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                shootAngleIntervalVals.Add((short)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "shootAngleXInterval")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                shootAngleXIntervalVals.Add((short)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "damageDamp")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                damageDampVals.Add((sbyte)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "spelDamageDamp")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                spelDamageDampVals.Add((sbyte)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "fireDamageDamp")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                fireDamageDampVals.Add((sbyte)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "thunderDamageDamp")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                thunderDamageDampVals.Add((sbyte)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "staminaDamp")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                staminaDampVals.Add((sbyte)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "knockbackDamp")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                knockbackDampVals.Add((sbyte)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "shootAngleXZ")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                shootAngleXZVals.Add((sbyte)(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "lockShootLimitAng")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                lockShootLimitAngVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "isPenetrate")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                isPenetrateVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "atkAttribute")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                atkAttributeVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "spAttribute")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                spAttributeVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "Material_AttackType")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                Material_AttackTypeVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "Material_AttackMaterial")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                Material_AttackMaterialVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "Material_Size")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                Material_SizeVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "launchConditionType")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                launchConditionTypeVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "FollowType")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                FollowTypeVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "isAttackSFX")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                isAttackSFXVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "isEndlessHit")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                isEndlessHitVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "isPenetrateMap")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                isPenetrateMapVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "isHitBothTeam")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                isHitBothTeamVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "isUseSharedHitList")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                isUseSharedHitListVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "isHitForceMagic")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                isHitForceMagicVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "isIgnoreSfxIfHitWater")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                isIgnoreSfxIfHitWaterVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "IsIgnoreMoveStateIfHitWater")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                IsIgnoreMoveStateIfHitWaterVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "isHitDarkForceMagic")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                isHitDarkForceMagicVals.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                        }
                    }

                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "atkId_Bullet")
                            {
                                int randomIndex = r.Next(atkId_BulletVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");
                                if (Convert.ToInt32(prop.GetValue(cell, null)) > 0)
                                {
                                    if (chkBullets.Checked)
                                    {
                                        prop.SetValue(cell, atkId_BulletVals[randomIndex], null);
                                    }
                                    atkId_BulletVals.RemoveAt(randomIndex);
                                }
                            }
                            else if (cell.Def.Name == "sfxId_Bullet")
                            {
                                int randomIndex = r.Next(sfxId_BulletVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");
                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, sfxId_BulletVals[randomIndex], null);
                                }
                                sfxId_BulletVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "sfxId_Hit")
                            {
                                int randomIndex = r.Next(sfxId_HitVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, sfxId_HitVals[randomIndex], null);
                                }

                                sfxId_HitVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "sfxId_Flick")
                            {
                                int randomIndex = r.Next(sfxId_FlickVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, sfxId_FlickVals[randomIndex], null);
                                }

                                sfxId_FlickVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "life")
                            {
                                int randomIndex = r.Next(lifeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, lifeVals[randomIndex], null);
                                }

                                lifeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "dist")
                            {
                                int randomIndex = r.Next(distVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if ((float)(prop.GetValue(cell, null)) > 0)
                                {
                                    if (chkBullets.Checked)
                                    {
                                        prop.SetValue(cell, distVals[randomIndex], null);
                                    }

                                    distVals.RemoveAt(randomIndex);
                                }
                            }
                            else if (cell.Def.Name == "shootInterval")
                            {
                                int randomIndex = r.Next(shootIntervalVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");
                                if ((float)(prop.GetValue(cell, null)) > 0)
                                {
                                    if (chkBullets.Checked)
                                    {
                                        prop.SetValue(cell, shootIntervalVals[randomIndex], null);
                                    }

                                    shootIntervalVals.RemoveAt(randomIndex);
                                }
                            }
                            else if (cell.Def.Name == "gravityInRange")
                            {
                                int randomIndex = r.Next(gravityInRangeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, gravityInRangeVals[randomIndex], null);
                                }

                                gravityInRangeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "gravityOutRange")
                            {
                                int randomIndex = r.Next(gravityOutRangeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, gravityOutRangeVals[randomIndex], null);
                                }

                                gravityOutRangeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "hormingStopRange")
                            {
                                int randomIndex = r.Next(hormingStopRangeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, hormingStopRangeVals[randomIndex], null);
                                }

                                hormingStopRangeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "initVellocity")
                            {
                                int randomIndex = r.Next(initVellocityVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");
                                if ((float)(prop.GetValue(cell, null)) > 0)
                                {
                                    if (chkBullets.Checked)
                                    {
                                        prop.SetValue(cell, initVellocityVals[randomIndex], null);
                                    }

                                    initVellocityVals.RemoveAt(randomIndex);
                                }
                            }
                            else if (cell.Def.Name == "accelInRange")
                            {
                                int randomIndex = r.Next(accelInRangeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, accelInRangeVals[randomIndex], null);
                                }

                                accelInRangeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "accelOutRange")
                            {
                                int randomIndex = r.Next(accelOutRangeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, accelOutRangeVals[randomIndex], null);
                                }

                                accelOutRangeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "maxVellocity")
                            {
                                int randomIndex = r.Next(maxVellocityVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");
                                if ((float)(prop.GetValue(cell, null)) > 0)
                                {
                                    if (chkBullets.Checked)
                                    {
                                        prop.SetValue(cell, maxVellocityVals[randomIndex], null);
                                    }

                                    maxVellocityVals.RemoveAt(randomIndex);
                                }
                            }
                            else if (cell.Def.Name == "minVellocity")
                            {
                                int randomIndex = r.Next(minVellocityVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, minVellocityVals[randomIndex], null);
                                }

                                minVellocityVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "accelTime")
                            {
                                int randomIndex = r.Next(accelTimeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, accelTimeVals[randomIndex], null);
                                }

                                accelTimeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "homingBeginDist")
                            {
                                int randomIndex = r.Next(homingBeginDistVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, homingBeginDistVals[randomIndex], null);
                                }

                                homingBeginDistVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "hitRadius")
                            {
                                int randomIndex = r.Next(hitRadiusVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, hitRadiusVals[randomIndex], null);
                                }

                                hitRadiusVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "hitRadiusMax")
                            {
                                int randomIndex = r.Next(hitRadiusMaxVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, hitRadiusMaxVals[randomIndex], null);
                                }

                                hitRadiusMaxVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "spreadTime")
                            {
                                int randomIndex = r.Next(spreadTimeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, spreadTimeVals[randomIndex], null);
                                }

                                spreadTimeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "hormingOffsetRange")
                            {
                                int randomIndex = r.Next(hormingOffsetRangeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, hormingOffsetRangeVals[randomIndex], null);
                                }

                                hormingOffsetRangeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "dmgHitRecordLifeTime")
                            {
                                int randomIndex = r.Next(dmgHitRecordLifeTimeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, dmgHitRecordLifeTimeVals[randomIndex], null);
                                }

                                dmgHitRecordLifeTimeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "spEffectIDForShooter")
                            {
                                int randomIndex = r.Next(spEffectIDForShooterVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, spEffectIDForShooterVals[randomIndex], null);
                                }

                                spEffectIDForShooterVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "HitBulletID")
                            {
                                int randomIndex = r.Next(HitBulletIDVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, HitBulletIDVals[randomIndex], null);
                                }

                                HitBulletIDVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "spEffectId0")
                            {
                                int randomIndex = r.Next(spEffectId0Vals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, spEffectId0Vals[randomIndex], null);
                                }

                                spEffectId0Vals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "numShoot")
                            {
                                int randomIndex = r.Next(numShootVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, numShootVals[randomIndex], null);
                                }

                                numShootVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "homingAngle")
                            {
                                int randomIndex = r.Next(homingAngleVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, homingAngleVals[randomIndex], null);
                                }

                                homingAngleVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "shootAngle")
                            {
                                int randomIndex = r.Next(shootAngleVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, shootAngleVals[randomIndex], null);
                                }

                                shootAngleVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "shootAngleInterval")
                            {
                                int randomIndex = r.Next(shootAngleIntervalVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, shootAngleIntervalVals[randomIndex], null);
                                }

                                shootAngleIntervalVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "shootAngleXInterval")
                            {
                                int randomIndex = r.Next(shootAngleXIntervalVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, shootAngleXIntervalVals[randomIndex], null);
                                }

                                shootAngleXIntervalVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "damageDamp")
                            {
                                int randomIndex = r.Next(damageDampVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, damageDampVals[randomIndex], null);
                                }

                                damageDampVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "spelDamageDamp")
                            {
                                int randomIndex = r.Next(spelDamageDampVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, spelDamageDampVals[randomIndex], null);
                                }

                                spelDamageDampVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "fireDamageDamp")
                            {
                                int randomIndex = r.Next(fireDamageDampVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, fireDamageDampVals[randomIndex], null);
                                }

                                fireDamageDampVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "thunderDamageDamp")
                            {
                                int randomIndex = r.Next(thunderDamageDampVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, thunderDamageDampVals[randomIndex], null);
                                }

                                thunderDamageDampVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "staminaDamp")
                            {
                                int randomIndex = r.Next(staminaDampVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, staminaDampVals[randomIndex], null);
                                }

                                staminaDampVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "knockbackDamp")
                            {
                                int randomIndex = r.Next(knockbackDampVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, knockbackDampVals[randomIndex], null);
                                }

                                knockbackDampVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "shootAngleXZ")
                            {
                                int randomIndex = r.Next(shootAngleXZVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, shootAngleXZVals[randomIndex], null);
                                }

                                shootAngleXZVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "lockShootLimitAng")
                            {
                                int randomIndex = r.Next(lockShootLimitAngVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, lockShootLimitAngVals[randomIndex], null);
                                }

                                lockShootLimitAngVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "isPenetrate")
                            {
                                int randomIndex = r.Next(isPenetrateVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, isPenetrateVals[randomIndex], null);
                                }

                                isPenetrateVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "atkAttribute")
                            {
                                int randomIndex = r.Next(atkAttributeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, atkAttributeVals[randomIndex], null);
                                }

                                atkAttributeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "spAttribute")
                            {
                                int randomIndex = r.Next(spAttributeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, spAttributeVals[randomIndex], null);
                                }

                                spAttributeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "Material_AttackType")
                            {
                                int randomIndex = r.Next(Material_AttackTypeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, Material_AttackTypeVals[randomIndex], null);
                                }

                                Material_AttackTypeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "Material_AttackMaterial")
                            {
                                int randomIndex = r.Next(Material_AttackMaterialVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, Material_AttackMaterialVals[randomIndex], null);
                                }

                                Material_AttackMaterialVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "Material_Size")
                            {
                                int randomIndex = r.Next(Material_SizeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, Material_SizeVals[randomIndex], null);
                                }

                                Material_SizeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "launchConditionType")
                            {
                                int randomIndex = r.Next(launchConditionTypeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, launchConditionTypeVals[randomIndex], null);
                                }

                                launchConditionTypeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "FollowType")
                            {
                                int randomIndex = r.Next(FollowTypeVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, FollowTypeVals[randomIndex], null);
                                }

                                FollowTypeVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "isAttackSFX")
                            {
                                int randomIndex = r.Next(isAttackSFXVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, isAttackSFXVals[randomIndex], null);
                                }

                                isAttackSFXVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "isEndlessHit")
                            {
                                int randomIndex = r.Next(isEndlessHitVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, isEndlessHitVals[randomIndex], null);
                                }

                                isEndlessHitVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "isPenetrateMap")
                            {
                                int randomIndex = r.Next(isPenetrateMapVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, isPenetrateMapVals[randomIndex], null);
                                }

                                isPenetrateMapVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "isHitBothTeam")
                            {
                                int randomIndex = r.Next(isHitBothTeamVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, isHitBothTeamVals[randomIndex], null);
                                }

                                isHitBothTeamVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "isUseSharedHitList")
                            {
                                int randomIndex = r.Next(isUseSharedHitListVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, isUseSharedHitListVals[randomIndex], null);
                                }

                                isUseSharedHitListVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "isHitForceMagic")
                            {
                                int randomIndex = r.Next(isHitForceMagicVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, isHitForceMagicVals[randomIndex], null);
                                }

                                isHitForceMagicVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "isIgnoreSfxIfHitWater")
                            {
                                int randomIndex = r.Next(isIgnoreSfxIfHitWaterVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, isIgnoreSfxIfHitWaterVals[randomIndex], null);
                                }

                                isIgnoreSfxIfHitWaterVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "IsIgnoreMoveStateIfHitWater")
                            {
                                int randomIndex = r.Next(IsIgnoreMoveStateIfHitWaterVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, IsIgnoreMoveStateIfHitWaterVals[randomIndex], null);
                                }

                                IsIgnoreMoveStateIfHitWaterVals.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "isHitDarkForceMagicList")
                            {
                                int randomIndex = r.Next(isHitDarkForceMagicVals.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkBullets.Checked)
                                {
                                    prop.SetValue(cell, isHitDarkForceMagicVals[randomIndex], null);
                                }

                                isHitDarkForceMagicVals.RemoveAt(randomIndex);
                            }
                        }
                    }
                }
            }

            //repack param files
            foreach (BND paramBnd in gameparamBnds)
            {
                foreach (MeowDSIO.DataTypes.BND.BNDEntry param in paramBnd)
                {
                    string filteredParamName = param.Name.Substring(param.Name.LastIndexOf("\\") + 1).Replace(".param", "");

                    PARAM matchingParam = AllParams.Where(x => x.VirtualUri == param.Name).First();

                    param.ReplaceData(matchingParam,new Progress<(int, int)>((p) =>
                    {

                    }));
                }

                DataFile.Resave(paramBnd, new Progress<(int, int)>((p) =>
                {

                }));
            }

            lblMessage.Text += "Randomizing Complete!";
            lblMessage.ForeColor = Color.Black;
            lblMessage.Visible = true;
        }

        class UiThread
        {
            public static void WriteToInfoLabel(IProgress<string> progress)
            {
                //why is this necessary
                //without the loop it doesnt run async
                for (var i = 0; i < 5; i++)
                {
                    Task.Delay(10).Wait();
                    progress.Report("Randomizing...\n\n");
                }
            }
        }

        public static string[] InvalidVoiceIds
        {
            //this is a list of all empty voice lines
            //nice job from
            get
            {
                string[] invalidIds = { "10010601", "10010602", "10010603", "10010604", "10010605", "10010606", "10010607", "10010608", "10010609", "10010611", "10010612", "10010613", "10010614",
                    "10010615", "10010616", "10010617", "10010618", "10010619", "10010621", "10010622", "10010623", "10010624", "10010625", "10010626", "10010627", "10010628", "10010629", "10010631",
                    "10010632", "10010633", "10010634", "10010635", "10010636", "10010637", "10010638", "10010639", "11000501", "11000502", "11000503", "11000504", "11000505", "11000506", "11000507",
                    "11000508", "11000509", "11000511", "11000512", "11000513", "11000514", "11000515", "11000516", "11000517", "11000518", "11000519", "11000521", "11000522", "11000523", "11000524",
                    "11000525", "11000526", "11000527", "11000528", "11000529", "11000601", "11000602", "11000603", "11000604", "11000605", "11000606", "11000607", "11000608", "11000609", "11000701",
                    "11000702", "11000703", "11000704", "11000705", "11000706", "11000707", "11000708", "11000709", "11000711", "11000712", "11000713", "11000714", "11000715", "11000716", "11000717",
                    "11000718", "11000719", "11000721", "11000722", "11000723", "11000724", "11000725", "11000726", "11000727", "11000728", "11000729", "13010402", "14000024", "16000000", "16000001",
                    "16000002", "16000003", "16000004", "16000005", "16000006", "16000007", "16000008", "16000009", "16000010", "16000011", "16000012", "16000013", "16000014", "16000015", "16000016",
                    "16000017", "16000018", "16000019", "16000020", "16000021", "16000022", "16000023", "16000024", "16000025", "16000026", "16000027", "16000028", "16000029", "16000030", "16000031",
                    "16000032", "16000033", "16000034", "16000035", "16000036", "16000037", "16000038", "16000039", "16000040", "16000041", "16000042", "16000043", "16000044", "16000045", "16000046",
                    "16000047", "16000048", "16000049", "16000050", "16000051", "16000052", "16000053", "16000054", "16000055", "16000056", "16000057", "16000058", "16000059", "16000060", "16000061",
                    "16000062", "16000063", "16000064", "16000065", "16000066", "16000067", "16000068", "16000069", "16000070", "16000071", "16000072", "16000073", "16000074", "16000075", "16000076",
                    "16000077", "16000078", "16000079", "16000080", "16000081", "16000082", "16000083", "16000084", "16000085", "16000086", "16000087", "16000088", "16000089", "16000090", "16000091",
                    "16000092", "16000093", "16000094", "16000095", "16000096", "16000097", "16000098", "16000099", "16000401", "16000402", "16000403", "16000404", "16000405", "16000406", "16000407",
                    "16000408", "16000409", "16000411", "16000412", "16000413", "16000414", "16000415", "16000416", "16000417", "16000418", "16000419", "16000421", "16000422", "16000423", "16000424",
                    "16000425", "16000426", "16000427", "16000428", "16000429", "16000431", "16000432", "16000433", "16000434", "16000435", "16000436", "16000437", "16000438", "16000439", "16000441",
                    "16000442", "16000443", "16000444", "16000445", "16000446", "16000447", "16000448", "16000449", "17000231", "17000232", "17000233", "17000234", "17000235", "17000236", "17000237",
                    "17000238", "17000239", "17000241", "17000242", "17000243", "17000244", "17000245", "17000246", "17000247", "17000248", "17000249", "18000401", "18000402", "18000403", "18000404",
                    "18000405", "18000406", "18000407", "18000408", "18000409", "18000411", "18000412", "18000413", "18000414", "18000415", "18000416", "18000417", "18000418", "18000419", "18000421",
                    "18000422", "18000423", "18000424", "18000425", "18000426", "18000427", "18000428", "18000429", "18010701", "18010702", "18010703", "18010704", "18010705", "18010706", "18010707",
                    "18010708", "18010709", "18010711", "18010712", "18010713", "18010714", "18010715", "18010716", "18010717", "18010718", "18010719", "18010721", "18010722", "18010723", "18010724",
                    "18010725", "18010726", "18010727", "18010728", "18010729", "19000421", "19000422", "19000423", "19000424", "19000425", "19000426", "19000427", "19000428", "19000429", "26000301",
                    "26001501", "26001502", "26001503", "26001504", "26001505", "26001506", "26001507", "26001508", "26001509", "26001511", "26001512", "26001513", "26001514", "26001515", "26001516",
                    "26001517", "26001518", "26001519", "26001521", "26001522", "26001523", "26001524", "26001525", "26001526", "26001527", "26001528", "26001529", "27000341", "28001001", "29001454",
                    "29001455", "29001456", "29001457", "29001458", "29001459", "29001464", "29001465", "29001466", "29001467", "29001468", "29001469", "32000026", "32000027", "32000028", "32000029",
                    "33000601", "33000602", "33000603", "33000604", "33000605", "33000606", "33000607", "33000608", "33000609", "33000611", "33000612", "33000613", "33000614", "33000615", "33000616",
                    "33000617", "33000618", "33000619", "33000621", "33000622", "33000623", "33000624", "33000625", "33000626", "33000627", "33000628", "33000629", "34000601", "34000602", "34000603",
                    "34000604", "34000605", "34000606", "34000607", "34000608", "34000609", "34000611", "34000612", "34000613", "34000614", "34000615", "34000616", "34000617", "34000618", "34000619",
                    "34000621", "34000622", "34000623", "34000624", "34000625", "34000626", "34000627", "34000628", "34000629", "35000100", "35000200", "35000220", "36000601", "36000602", "36000603",
                    "36000604", "36000605", "36000606", "36000607", "36000608", "36000609", "36000611", "36000612", "36000613", "36000614", "36000615", "36000616", "36000617", "36000618", "36000619",
                    "36000621", "36000622", "36000623", "36000624", "36000625", "36000626", "36000627", "36000628", "36000629", "36000631", "36000632", "36000633", "36000634", "36000635", "36000636",
                    "36000637", "36000638", "36000639", "36000701", "36000702", "36000703", "36000704", "36000705", "36000706", "36000707", "36000708", "36000709", "36000711", "36000712", "36000713",
                    "36000714", "36000715", "36000716", "36000717", "36000718", "36000719", "36000801", "36000802", "36000803", "36000804", "36000805", "36000806", "36000807", "36000808", "36000809",
                    "37002701", "37002702", "37002703", "37002704", "37002705", "37002706", "37002707", "37002708", "37002709", "37002711", "37002712", "37002713", "37002714", "37002715", "37002716",
                    "37002717", "37002718", "37002719", "37002721", "37002722", "37002723", "37002724", "37002725", "37002726", "37002727", "37002728", "37002729", "38040105", "38060108", "38060200",
                    "38060300", "39000231", "39000232", "39000233", "39000234", "39000235", "39000236", "39000237", "39000238", "39000239", "40010101", "41000301", "41000302", "41000303", "41000304",
                    "41000305", "41000306", "41000307", "41000308", "41000309", "41000311", "41000312", "41000313", "41000314", "41000315", "41000316", "41000317", "41000318", "41000319", "41000321",
                    "41000322", "41000323", "41000324", "41000325", "41000326", "41000327", "41000328", "41000329", "42000841", "42000842", "42000843", "42000844", "42000845", "42000846", "42000847",
                    "42000848", "42000849", "42000851", "42000852", "42000853", "42000854", "42000855", "42000856", "42000857", "42000858", "42000859", "42001241", "42001242", "42001243", "42001244",
                    "42001245", "42001246", "42001247", "42001248", "42001249", "42001251", "42001252", "42001253", "42001254", "42001255", "42001256", "42001257", "42001258", "42001259", "43001004",
                    "43001209", "44000400", "44001205", "44001304", "44001402", "44001403", "44001404", "44001405", "47000901", "47000902", "47000903", "47000904", "47000905", "47000906", "47000907",
                    "47000908", "47000909", "47000911", "47000912", "47000913", "47000914", "47000915", "47000916", "47000917", "47000918", "47000919", "47000921", "47000922", "47000923", "47000924",
                    "47000925", "47000926", "47000927", "47000928", "47000929", "47000931", "47000932", "47000933", "47000934", "47000935", "47000936", "47000937", "47000938", "47000939", "52000000",
                    "52000100", "11000000", "14012502", "16000103", "28001908", "29000300", "29000301", "29000302", "29000303", "29000304", "29000305", "29000401", "29000602", "29002702", "36000502",
                    "37000703", "37000802", "37001301", "38000105", "39030101", "40010123", "42000205", "43001006", "43001202", "44001201", "44001202", "44001203", "44001307", "44002009", "56000001",
                    "56000002", "56000003", "56000004", "56000005", "56000006", "56000007", "56000008", "56000009", "56000011", "56000012", "56000013", "56000014", "56000015", "56000016", "56000017",
                    "56000018", "56000019", "56000101", "56000102", "56000103", "56000104", "56000105", "56000106", "56000107", "56000108", "56000109", "56000111", "56000112", "56000113", "56000114",
                    "56000115", "56000116", "56000117", "56000118", "56000119", "56000121", "56000122", "56000123", "56000124", "56000125", "56000126", "56000127", "56000128", "56000129", "56000131",
                    "56000132", "56000133", "56000134", "56000135", "56000136", "56000137", "56000138", "56000139", "56000141", "56000142", "56000143", "56000144", "56000145", "56000146", "56000147",
                    "56000148", "56000149", "56000151", "56000152", "56000153", "56000154", "56000155", "56000156", "56000157", "56000158", "56000159", "56000161", "56000162", "56000163", "56000164",
                    "56000165", "56000166", "56000167", "56000168", "56000169", "56000170", "56000171", "56000172", "56000173", "56000174", "56000175", "56000176", "56000177", "56000178", "56000179",
                    "56000181", "56000182", "56000183", "56000184", "56000185", "56000186", "56000187", "56000188", "56000189", "56000191", "56000192", "56000193", "56000194", "56000195", "56000196",
                    "56000197", "56000198", "56000199", "56000201", "56000202", "56000203", "56000204", "56000205", "56000206", "56000207", "56000208", "56000209", "56000211", "56000212", "56000213",
                    "56000214", "56000215", "56000216", "56000217", "56000218", "56000219", "56000221", "56000222", "56000223", "56000224", "56000225", "56000226", "56000227", "56000228", "56000229",
                    "56000301", "56000302", "56000303", "56000304", "56000305", "56000306", "56000307", "56000308", "56000309", "56000311", "56000312", "56000313", "56000314", "56000315", "56000316",
                    "56000317", "56000318", "56000319", "56000321", "56000322", "56000323", "56000324", "56000325", "56000326", "56000327", "56000328", "56000329", "56000331", "56000332", "56000333",
                    "56000334", "56000335", "56000336", "56000337", "56000338", "56000339", "56000341", "56000342", "56000343", "56000344", "56000345", "56000346", "56000347", "56000348", "56000349",
                    "56000351", "56000352", "56000353", "56000354", "56000355", "56000356", "56000357", "56000358", "56000359", "56000361", "56000362", "56000363", "56000364", "56000365", "56000366",
                    "56000367", "56000368", "56000369", "56005001", "56005002", "56005003", "56005004", "56005005", "56005006", "56005007", "56005008", "56005009", "57000001", "57000002", "57000003",
                    "57000004", "57000005", "57000006", "57000007", "57000008", "57000009", "57000011", "57000012", "57000013", "57000014", "57000015", "57000016", "57000017", "57000018", "57000019",
                    "57000021", "57000022", "57000023", "57000024", "57000025", "57000026", "57000027", "57000028", "57000029", "57000031", "57000032", "57000033", "57000034", "57000035", "57000036",
                    "57000037", "57000038", "57000039", "57000041", "57000042", "57000043", "57000044", "57000045", "57000046", "57000047", "57000048", "57000049", "57000101", "57000102", "57000103",
                    "57000104", "57000105", "57000106", "57000107", "57000108", "57000109", "57000111", "57000112", "57000113", "57000114", "57000115", "57000116", "57000117", "57000118", "57000119",
                    "57000121", "57000122", "57000123", "57000124", "57000125", "57000126", "57000127", "57000128", "57000129", "57000131", "57000132", "57000133", "57000134", "57000135", "57000136",
                    "57000137", "57000138", "57000139", "57000141", "57000142", "57000143", "57000144", "57000145", "57000146", "57000147", "57000148", "57000149", "57005001", "57005002", "57005003",
                    "57005004", "57005005", "57005006", "57005007", "57005008", "57005009", "57005011", "57005012", "57005013", "57005014", "57005015", "57005016", "57005017", "57005018", "57005019",
                    "57005101", "57005102", "57005103", "57005104", "57005105", "57005106", "57005107", "57005108", "57005109", "57005111", "57005112", "57005113", "57005114", "57005115", "57005116",
                    "57005117", "57005118", "57005119", "57005201", "57005202", "57005203", "57005204", "57005205", "57005206", "57005207", "57005208", "57005209", "57005211", "57005212", "57005213",
                    "57005214", "57005215", "57005216", "57005217", "57005218", "57005219", "58000001", "58000002", "58000003", "58000004", "58000005", "58000006", "58000007", "58000008", "58000009",
                    "58000011", "58000012", "58000013", "58000014", "58000015", "58000016", "58000017", "58000018", "58000019", "58000021", "58000022", "58000023", "58000024", "58000025", "58000026",
                    "58000027", "58000028", "58000029", "58000031", "58000032", "58000033", "58000034", "58000035", "58000036", "58000037", "58000038", "58000039", "58000041", "58000042", "58000043",
                    "58000044", "58000045", "58000046", "58000047", "58000048", "58000049", "58000401", "58000402", "58000403", "58000404", "58000405", "58000406", "58000407", "58000408", "58000409",
                    "58000411", "58000412", "58000413", "58000414", "58000415", "58000416", "58000417", "58000418", "58000419", "58000421", "58000422", "58000423", "58000424", "58000425", "58000426",
                    "58000427", "58000428", "58000429", "58000431", "58000432", "58000433", "58000434", "58000435", "58000436", "58000437", "58000438", "58000439", "58000501", "58000502", "58000503",
                    "58000504", "58000505", "58000506", "58000507", "58000508", "58000509", "58000511", "58000512", "58000513", "58000514", "58000515", "58000516", "58000517", "58000518", "58000519",
                    "58000521", "58000522", "58000523", "58000524", "58000525", "58000526", "58000527", "58000528", "58000529", "58000531", "58000532", "58000533", "58000534", "58000535", "58000536",
                    "58000537", "58000538", "58000539", "58000541", "58000542", "58000543", "58000544", "58000545", "58000546", "58000547", "58000548", "58000549", "58000551", "58000552", "58000553",
                    "58000554", "58000555", "58000556", "58000557", "58000558", "58000559", "58000561", "58000562", "58000563", "58000564", "58000565", "58000566", "58000567", "58000568", "58000569",
                    "58000571", "58000572", "58000573", "58000574", "58000575", "58000576", "58000577", "58000578", "58000579", "58000581", "58000582", "58000583", "58000584", "58000585", "58000586",
                    "58000587", "58000588", "58000589", "58000591", "58000592", "58000593", "58000594", "58000595", "58000596", "58000597", "58000598", "58000599", "58001201", "58001202", "58001203",
                    "58001204", "58001205", "58001206", "58001207", "58001208", "58001209", "58001301", "58001302", "58001303", "58001304", "58001305", "58001306", "58001307", "58001308", "58001309",
                    "58001311", "58001312", "58001313", "58001314", "58001315", "58001316", "58001317", "58001318", "58001319", "58001401", "58001402", "58001403", "58001404", "58001405", "58001406",
                    "58001407", "58001408", "58001409", "58001501", "58001502", "58001503", "58001504", "58001505", "58001506", "58001507", "58001508", "58001509", "58001601", "58001602", "58001603",
                    "58001604", "58001605", "58001606", "58001607", "58001608", "58001609", "58001701", "58001702", "58001703", "58001704", "58001705", "58001706", "58001707", "58001708", "58001709",
                    "58001801", "58001802", "58001803", "58001804", "58001805", "58001806", "58001807", "58001808", "58001809", "58001811", "58001812", "58001813", "58001814", "58001815", "58001816",
                    "58001817", "58001818", "58001819", "58001821", "58001822", "58001823", "58001824", "58001825", "58001826", "58001827", "58001828", "58001829", "58001901", "58001902", "58001903",
                    "58001904", "58001905", "58001906", "58001907", "58001908", "58001909", "58001911", "58001912", "58001913", "58001914", "58001915", "58001916", "58001917", "58001918", "58001919",
                    "58001921", "58001922", "58001923", "58001924", "58001925", "58001926", "58001927", "58001928", "58001929", "58002001", "58002002", "58002003", "58002004", "58002005", "58002006",
                    "58002007", "58002008", "58002009", "58002011", "58002012", "58002013", "58002014", "58002015", "58002016", "58002017", "58002018", "58002019", "58002021", "58002022", "58002023",
                    "58002024", "58002025", "58002026", "58002027", "58002028", "58002029", "58002031", "58002032", "58002033", "58002034", "58002035", "58002036", "58002037", "58002038", "58002039",
                    "58002101", "58002102", "58002103", "58002104", "58002105", "58002106", "58002107", "58002108", "58002109", "58002111", "58002112", "58002113", "58002114", "58002115", "58002116",
                    "58002117", "58002118", "58002119", "58002121", "58002122", "58002123", "58002124", "58002125", "58002126", "58002127", "58002128", "58002129", "58002131", "58002132", "58002133",
                    "58002134", "58002135", "58002136", "58002137", "58002138", "58002139", "58002201", "58002202", "58002203", "58002204", "58002205", "58002206", "58002207", "58002208", "58002209",
                    "58002211", "58002212", "58002213", "58002214", "58002215", "58002216", "58002217", "58002218", "58002219", "58002221", "58002222", "58002223", "58002224", "58002225", "58002226",
                    "58002227", "58002228", "58002229", "58002301", "58002302", "58002303", "58002304", "58002305", "58002306", "58002307", "58002308", "58002309", "58002311", "58002312", "58002313",
                    "58002314", "58002315", "58002316", "58002317", "58002318", "58002319", "58002321", "58002322", "58002323", "58002324", "58002325", "58002326", "58002327", "58002328", "58002329",
                    "58002331", "58002332", "58002333", "58002334", "58002335", "58002336", "58002337", "58002338", "58002339", "58002341", "58002342", "58002343", "58002344", "58002345", "58002346",
                    "58002347", "58002348", "58002349", "58002351", "58002352", "58002353", "58002354", "58002355", "58002356", "58002357", "58002358", "58002359", "58002361", "58002362", "58002363",
                    "58002364", "58002365", "58002366", "58002367", "58002368", "58002369", "58002401", "58002402", "58002403", "58002404", "58002405", "58002406", "58002407", "58002408", "58002409",
                    "58002411", "58002412", "58002413", "58002414", "58002415", "58002416", "58002417", "58002418", "58002419", "58002421", "58002422", "58002423", "58002424", "58002425", "58002426",
                    "58002427", "58002428", "58002429", "58002501", "58002502", "58002503", "58002504", "58002505", "58002506", "58002507", "58002508", "58002509", "58002601", "58002602", "58002603",
                    "58002604", "58002605", "58002606", "58002607", "58002608", "58002609", "58002611", "58002612", "58002613", "58002614", "58002615", "58002616", "58002617", "58002618", "58002619",
                    "58002621", "58002622", "58002623", "58002624", "58002625", "58002626", "58002627", "58002628", "58002629", "58002631", "58002632", "58002633", "58002634", "58002635", "58002636",
                    "58002637", "58002638", "58002639", "58002641", "58002642", "58002643", "58002644", "58002645", "58002646", "58002647", "58002648", "58002649", "58002701", "58002702", "58002703",
                    "58002704", "58002705", "58002706", "58002707", "58002708", "58002709", "58002711", "58002712", "58002713", "58002714", "58002715", "58002716", "58002717", "58002718", "58002719",
                    "58002721", "58002722", "58002723", "58002724", "58002725", "58002726", "58002727", "58002728", "58002729", "58002731", "58002732", "58002733", "58002734", "58002735", "58002736",
                    "58002737", "58002738", "58002739", "58002741", "58002742", "58002743", "58002744", "58002745", "58002746", "58002747", "58002748", "58002749", "58005001", "58005002", "58005003",
                    "58005004", "58005005", "58005006", "58005007", "58005008", "58005009", "58005011", "58005012", "58005013", "58005014", "58005015", "58005016", "58005017", "58005018", "58005019",
                    "58005021", "58005022", "58005023", "58005024", "58005025", "58005026", "58005027", "58005028", "58005029", "58005031", "58005032", "58005033", "58005034", "58005035", "58005036",
                    "58005037", "58005038", "58005039", "58005101", "58005102", "58005103", "58005104", "58005105", "58005106", "58005107", "58005108", "58005109", "58005111", "58005112", "58005113",
                    "58005114", "58005115", "58005116", "58005117", "58005118", "58005119", "58005201", "58005202", "58005203", "58005204", "58005205", "58005206", "58005207", "58005208", "58005209",
                    "58005211", "58005212", "58005213", "58005214", "58005215", "58005216", "58005217", "58005218", "58005219", "58005301", "58005302", "58005303", "58005304", "58005305", "58005306",
                    "58005307", "58005308", "58005309", "58005311", "58005312", "58005313", "58005314", "58005315", "58005316", "58005317", "58005318", "58005319", "58005401", "58005402", "58005403",
                    "58005404", "58005405", "58005406", "58005407", "58005408", "58005409", "58005411", "58005412", "58005413", "58005414", "58005415", "58005416", "58005417", "58005418", "58005419",
                    "59000001", "59000002", "59000003", "59000004", "59000005", "59000006", "59000007", "59000008", "59000009", "59000101", "59000102", "59000103", "59000104", "59000105", "59000106",
                    "59000107", "59000108", "59000109", "59000111", "59000112", "59000113", "59000114", "59000115", "59000116", "59000117", "59000118", "59000119", "59000201", "59000202", "59000203",
                    "59000204", "59000205", "59000206", "59000207", "59000208", "59000209", "59000211", "59000212", "59000213", "59000214", "59000215", "59000216", "59000217", "59000218", "59000219",
                    "59000301", "59000302", "59000303", "59000304", "59000305", "59000306", "59000307", "59000308", "59000309", "59000401", "59000402", "59000403", "59000404", "59000405", "59000406",
                    "59000407", "59000408", "59000409", "59000411", "59000412", "59000413", "59000414", "59000415", "59000416", "59000417", "59000418", "59000419", "59000501", "59000502", "59000503",
                    "59000504", "59000505", "59000506", "59000507", "59000508", "59000509", "59000511", "59000512", "59000513", "59000514", "59000515", "59000516", "59000517", "59000518", "59000519",
                    "59000701", "59000702", "59000703", "59000704", "59000705", "59000706", "59000707", "59000708", "59000709", "59000711", "59000712", "59000713", "59000714", "59000715", "59000716",
                    "59000717", "59000718", "59000719", "59000801", "59000802", "59000803", "59000804", "59000805", "59000806", "59000807", "59000808", "59000809", "59000811", "59000812", "59000813",
                    "59000814", "59000815", "59000816", "59000817", "59000818", "59000819", "59000821", "59000822", "59000823", "59000824", "59000825", "59000826", "59000827", "59000828", "59000829",
                    "59000831", "59000832", "59000833", "59000834", "59000835", "59000836", "59000837", "59000838", "59000839", "59000901", "59000902", "59000903", "59000904", "59000905", "59000906",
                    "59000907", "59000908", "59000909", "59000911", "59000912", "59000913", "59000914", "59000915", "59000916", "59000917", "59000918", "59000919", "59000921", "59000922", "59000923",
                    "59000924", "59000925", "59000926", "59000927", "59000928", "59000929", "59000931", "59000932", "59000933", "59000934", "59000935", "59000936", "59000937", "59000938", "59000939",
                    "59001001", "59001002", "59001003", "59001004", "59001005", "59001006", "59001007", "59001008", "59001009", "59001011", "59001012", "59001013", "59001014", "59001015", "59001016",
                    "59001017", "59001018", "59001019", "59001021", "59001022", "59001023", "59001024", "59001025", "59001026", "59001027", "59001028", "59001029", "59001101", "59001102", "59001103",
                    "59001104", "59001105", "59001106", "59001107", "59001108", "59001109", "59001111", "59001112", "59001113", "59001114", "59001115", "59001116", "59001117", "59001118", "59001119",
                    "59001201", "59001202", "59001203", "59001204", "59001205", "59001206", "59001207", "59001208", "59001209", "59001211", "59001212", "59001213", "59001214", "59001215", "59001216",
                    "59001217", "59001218", "59001219", "59001301", "59001302", "59001303", "59001304", "59001305", "59001306", "59001307", "59001308", "59001309", "59001311", "59001312", "59001313",
                    "59001314", "59001315", "59001316", "59001317", "59001318", "59001319", "59001401", "59001402", "59001403", "59001404", "59001405", "59001406", "59001407", "59001408", "59001409",
                    "59001501", "59001502", "59001503", "59001504", "59001505", "59001506", "59001507", "59001508", "59001509", "59001601", "59001602", "59001603", "59001604", "59001605", "59001606",
                    "59001607", "59001608", "59001609", "59001611", "59001612", "59001613", "59001614", "59001615", "59001616", "59001617", "59001618", "59001619", "59001621", "59001622", "59001623",
                    "59001624", "59001625", "59001626", "59001627", "59001628", "59001629", "59001631", "59001632", "59001633", "59001634", "59001635", "59001636", "59001637", "59001638", "59001639",
                    "59001641", "59001642", "59001643", "59001644", "59001645", "59001646", "59001647", "59001648", "59001649", "59001651", "59001652", "59001653", "59001654", "59001655", "59001656",
                    "59001657", "59001658", "59001659", "59005001", "59005002", "59005003", "59005004", "59005005", "59005006", "59005007", "59005008", "59005009", "59005101", "59005102", "59005103",
                    "59005104", "59005105", "59005106", "59005107", "59005108", "59005109", "59005111", "59005112", "59005113", "59005114", "59005115", "59005116", "59005117", "59005118", "59005119",
                    "59005201", "59005202", "59005203", "59005204", "59005205", "59005206", "59005207", "59005208", "59005209", "59005211", "59005212", "59005213", "59005214", "59005215", "59005216",
                    "59005217", "59005218", "59005219", "59005221", "59005222", "59005223", "59005224", "59005225", "59005226", "59005227", "59005228", "59005229", "59005301", "59005302", "59005303",
                    "59005304", "59005305", "59005306", "59005307", "59005308", "59005309", "59005401", "59005402", "59005403", "59005404", "59005405", "59005406", "59005407", "59005408", "59005409",
                    "59005411", "59005412", "59005413", "59005414", "59005415", "59005416", "59005417", "59005418", "59005419", "59005421", "59005422", "59005423", "59005424", "59005425", "59005426",
                    "59005427", "59005428", "59005429", "59005501", "59005502", "59005503", "59005504", "59005505", "59005506", "59005507", "59005508", "59005509", "59005601", "59005602", "59005603",
                    "59005604", "59005605", "59005606", "59005607", "59005608", "59005609", "59005611", "59005612", "59005613", "59005614", "59005615", "59005616", "59005617", "59005618", "59005619",
                    "60000001", "60000002", "60000003", "60000004", "60000005", "60000006", "60000007", "60000008", "60000009", "60000012", "60000013", "60000014", "60000015", "60000016", "60000017",
                    "60000018", "60000019", "60000101", "60000102", "60000103", "60000104", "60000105", "60000106", "60000107", "60000108", "60000109", "60000111", "60000112", "60000113", "60000114",
                    "60000115", "60000116", "60000117", "60000118", "60000119", "60000121", "60000122", "60000123", "60000124", "60000125", "60000126", "60000127", "60000128", "60000129", "60000131",
                    "60000132", "60000133", "60000134", "60000135", "60000136", "60000137", "60000138", "60000139", "60000141", "60000142", "60000143", "60000144", "60000145", "60000146", "60000147",
                    "60000148", "60000149", "60000201", "60000202", "60000203", "60000204", "60000205", "60000206", "60000207", "60000208", "60000209", "60000211", "60000212", "60000213", "60000214",
                    "60000215", "60000216", "60000217", "60000218", "60000219", "60000301", "60000302", "60000303", "60000304", "60000305", "60000306", "60000307", "60000308", "60000309", "60000310",
                    "60000311", "60000312", "60000313", "60000314", "60000315", "60000316", "60000317", "60000318", "60000319", "60000321", "60000322", "60000323", "60000324", "60000325", "60000326",
                    "60000327", "60000328", "60000329", "60000401", "60000402", "60000403", "60000404", "60000405", "60000406", "60000407", "60000408", "60000409", "60000601", "60000602", "60000603",
                    "60000604", "60000605", "60000606", "60000607", "60000608", "60000609", "60000611", "60000612", "60000613", "60000614", "60000615", "60000616", "60000617", "60000618", "60000619",
                    "60000621", "60000622", "60000623", "60000624", "60000625", "60000626", "60000627", "60000628", "60000629", "60000701", "60000702", "60000703", "60000704", "60000705", "60000706",
                    "60000707", "60000708", "60000709", "60000711", "60000712", "60000713", "60000714", "60000715", "60000716", "60000717", "60000718", "60000719", "60000721", "60000722", "60000723",
                    "60000724", "60000725", "60000726", "60000727", "60000728", "60000729", "60005001", "60005002", "60005003", "60005004", "60005005", "60005006", "60005007", "60005008", "60005009",
                    "60005101", "60005102", "60005103", "60005104", "60005105", "60005106", "60005107", "60005108", "60005109", "60005111", "60005112", "60005113", "60005114", "60005115", "60005116",
                    "60005117", "60005118", "60005119", "60005201", "60005202", "60005203", "60005204", "60005205", "60005206", "60005207", "60005208", "60005209", "60005301", "60005302", "60005303",
                    "60005304", "60005305", "60005306", "60005307", "60005308", "60005309", "61000001", "61000002", "61000003", "61000004", "61000005", "61000006", "61000007", "61000008", "61000009",
                    "61000011", "61000012", "61000013", "61000014", "61000015", "61000016", "61000017", "61000018", "61000019", "61000021", "61000022", "61000023", "61000024", "61000025", "61000026",
                    "61000027", "61000028", "61000029", "61000032", "61000033", "61000034", "61000035", "61000036", "61000037", "61000038", "61000039", "61000041", "61000042", "61000043", "61000044",
                    "61000045", "61000046", "61000047", "61000048", "61000049", "61000051", "61000052", "61000053", "61000054", "61000055", "61000056", "61000057", "61000058", "61000059", "61000061",
                    "61000062", "61000063", "61000064", "61000065", "61000066", "61000067", "61000068", "61000069", "61000101", "61000102", "61000103", "61000104", "61000105", "61000106", "61000107",
                    "61000108", "61000109", "61000111", "61000112", "61000113", "61000114", "61000115", "61000116", "61000117", "61000118", "61000119", "61000121", "61000122", "61000123", "61000124",
                    "61000125", "61000126", "61000127", "61000128", "61000129", "61000131", "61000132", "61000133", "61000134", "61000135", "61000136", "61000137", "61000138", "61000139", "61000201",
                    "61000202", "61000203", "61000204", "61000205", "61000206", "61000207", "61000208", "61000209", "61000211", "61000212", "61000213", "61000214", "61000215", "61000216", "61000217",
                    "61000218", "61000219", "61000221", "61000222", "61000223", "61000224", "61000225", "61000226", "61000227", "61000228", "61000229", "61000231", "61000232", "61000233", "61000234",
                    "61000235", "61000236", "61000237", "61000238", "61000239", "61000241", "61000242", "61000243", "61000244", "61000245", "61000246", "61000247", "61000248", "61000249", "61000301",
                    "61000302", "61000303", "61000304", "61000305", "61000306", "61000307", "61000308", "61000309", "61000401", "61000402", "61000403", "61000404", "61000405", "61000406", "61000407",
                    "61000408", "61000409", "61000501", "61000502", "61000503", "61000504", "61000505", "61000506", "61000507", "61000508", "61000509", "61000601", "61000602", "61000603", "61000604",
                    "61000605", "61000606", "61000607", "61000608", "61000609", "61000701", "61000702", "61000703", "61000704", "61000705", "61000706", "61000707", "61000708", "61000709", "61000711",
                    "61000712", "61000713", "61000714", "61000715", "61000716", "61000717", "61000718", "61000719", "61000721", "61000722", "61000723", "61000724", "61000725", "61000726", "61000727",
                    "61000728", "61000729", "61000731", "61000732", "61000733", "61000734", "61000735", "61000736", "61000737", "61000738", "61000739", "61000901", "61000902", "61000903", "61000904",
                    "61000905", "61000906", "61000907", "61000908", "61000909", "61000911", "61000912", "61000913", "61000914", "61000915", "61000916", "61000917", "61000918", "61000919", "61000921",
                    "61000922", "61000923", "61000924", "61000925", "61000926", "61000927", "61000928", "61000929", "61000931", "61000932", "61000933", "61000934", "61000935", "61000936", "61000937",
                    "61000938", "61000939", "61000941", "61000942", "61000943", "61000944", "61000945", "61000946", "61000947", "61000948", "61000949", "61001001", "61001002", "61001003", "61001004",
                    "61001005", "61001006", "61001007", "61001008", "61001009", "61001011", "61001012", "61001013", "61001014", "61001015", "61001016", "61001017", "61001018", "61001019", "61001021",
                    "61001022", "61001023", "61001024", "61001025", "61001026", "61001027", "61001028", "61001029", "61001101", "61001102", "61001103", "61001104", "61001105", "61001106", "61001107",
                    "61001108", "61001109", "61001111", "61001112", "61001113", "61001114", "61001115", "61001116", "61001117", "61001118", "61001119", "61001121", "61001122", "61001123", "61001124",
                    "61001125", "61001126", "61001127", "61001128", "61001129", "61001201", "61001202", "61001203", "61001204", "61001205", "61001206", "61001207", "61001208", "61001209", "61001211",
                    "61001212", "61001213", "61001214", "61001215", "61001216", "61001217", "61001218", "61001219", "61001221", "61001222", "61001223", "61001224", "61001225", "61001226", "61001227",
                    "61001228", "61001229", "61001231", "61001232", "61001233", "61001234", "61001235", "61001236", "61001237", "61001238", "61001239", "61001301", "61001302", "61001303", "61001304",
                    "61001305", "61001306", "61001307", "61001308", "61001309", "61001311", "61001312", "61001313", "61001314", "61001315", "61001316", "61001317", "61001318", "61001319", "61001321",
                    "61001322", "61001323", "61001324", "61001325", "61001326", "61001327", "61001328", "61001329", "61005001", "61005002", "61005003", "61005004", "61005005", "61005006", "61005007",
                    "61005008", "61005009", "61005011", "61005012", "61005013", "61005014", "61005015", "61005016", "61005017", "61005018", "61005019", "61005021", "61005022", "61005023", "61005024",
                    "61005025", "61005026", "61005027", "61005028", "61005029", "61005031", "61005032", "61005033", "61005034", "61005035", "61005036", "61005037", "61005038", "61005039", "61005041",
                    "61005042", "61005043", "61005044", "61005045", "61005046", "61005047", "61005048", "61005049", "61005101", "61005102", "61005103", "61005104", "61005105", "61005106", "61005107",
                    "61005108", "61005109", "61005111", "61005112", "61005113", "61005114", "61005115", "61005116", "61005117", "61005118", "61005119", "61005121", "61005122", "61005123", "61005124",
                    "61005125", "61005126", "61005127", "61005128", "61005129" };
                return invalidIds;
            }
            
        }

    }
}
