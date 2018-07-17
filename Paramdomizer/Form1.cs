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
            gameDirectory = Assembly.GetEntryAssembly().Location;

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
            }
        }

        private async void btnSubmit_Click(object sender, EventArgs e)
        {
            //check that entered path is valid
            gameDirectory = txtGamePath.Text;

            //reset message label
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
            if (!File.Exists(gameDirectory + "\\param\\GameParam\\GameParam.parambndbak"))
            {
                File.Copy(gameDirectory + "\\param\\GameParam\\GameParam.parambnd", gameDirectory + "\\param\\GameParam\\GameParam.parambndbak");
                lblMessage.Text = "Backed up GameParam.parambnd at /DATA/param/GameParam/GameParam.parambndbak\n\n";
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
                            else if (cell.Def.Name == "stamina")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allStaminas.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "staminaRecoverBaseVel")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allStaminaRegens.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
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
                            else if (cell.Def.Name == "stamina")
                            {
                                int randomIndex = r.Next(allStaminas.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkStaminaRegen.Checked)
                                {
                                    prop.SetValue(cell, allStaminas[randomIndex], null);
                                }

                                allStaminas.RemoveAt(randomIndex);
                            }
                            else if (cell.Def.Name == "staminaRecoverBaseVal")
                            {
                                int randomIndex = r.Next(allStaminaRegens.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkStaminaRegen.Checked)
                                {
                                    prop.SetValue(cell, allStaminaRegens[randomIndex], null);
                                }

                                allStaminaRegens.RemoveAt(randomIndex);
                            }
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
                            if (cell.Def.Name == "nearDist")
                            {
                                int randomIndex = r.Next(allNearDists.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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

                                if (chkAggroRadius.Checked)
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
                            if (cell.Def.Name == "UseAnimation")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allUseAnimations.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                        }
                    }

                    //loop again to set a random value per entry
                    foreach (MeowDSIO.DataTypes.PARAM.ParamRow paramRow in paramFile.Entries)
                    {
                        foreach (MeowDSIO.DataTypes.PARAM.ParamCellValueRef cell in paramRow.Cells)
                        {
                            if (cell.Def.Name == "UseAnimation")
                            {
                                int randomIndex = r.Next(allUseAnimations.Count);
                                Type type = cell.GetType();
                                PropertyInfo prop = type.GetProperty("Value");

                                if (chkItemAnimations.Checked)
                                {
                                    prop.SetValue(cell, allUseAnimations[randomIndex], null);
                                }

                                allUseAnimations.RemoveAt(randomIndex);
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
                            if (cell.Def.Name == "SfxVariationId")
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
                            if (cell.Def.Name == "SfxVariationId")
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
                                allSounds.Add(Convert.ToInt32(prop.GetValue(cell, null)));
                            }
                            else if (cell.Def.Name == "msgId")
                            {
                                PropertyInfo prop = cell.GetType().GetProperty("Value");
                                allMsgs.Add(Convert.ToInt32(prop.GetValue(cell, null)));
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

                                if (chkVoices.Checked)
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
                                if (chkVoices.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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

                                if (chkSkeletons.Checked)
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
                for (var i = 0; i < 1; i++)
                {
                    Task.Delay(1).Wait();
                    progress.Report("Randomizing...\n\n");
                }
            }
        }

    }
}
