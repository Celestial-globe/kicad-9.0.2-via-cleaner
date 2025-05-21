#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import wx
import pcbnew
import time
import math

class ViaCleanerDialog(wx.Dialog):
    def __init__(self, parent):
        wx.Dialog.__init__(self, parent, title="VIA クリーナー", size=(450, 400))
        
        self.clearance = 0.2  # デフォルトクリアランス値 (mm)
        self.board_edge_clearance = 0.3  # 基板エッジクリアランス値 (mm)
        self.zone_clearance = 0.2  # ゾーンエッジクリアランス値 (mm)
        
        # メインパネルとスクロールバー
        main_panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # スクロール可能なパネル
        scroll_panel = wx.ScrolledWindow(main_panel, style=wx.VSCROLL)
        scroll_panel.SetScrollRate(0, 10)
        vbox = wx.BoxSizer(wx.VERTICAL)
        
        # パネルの内容をスクロールパネルに追加
        panel = wx.Panel(scroll_panel)
        panel_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # クリアランス設定
        clearance_box = wx.BoxSizer(wx.HORIZONTAL)
        clearance_label = wx.StaticText(panel, label="最小クリアランス (mm):", size=(200, -1))
        clearance_box.Add(clearance_label, flag=wx.RIGHT|wx.ALIGN_CENTER_VERTICAL, border=8)
        self.clearance_ctrl = wx.TextCtrl(panel, value=str(self.clearance))
        clearance_box.Add(self.clearance_ctrl, proportion=1)
        panel_sizer.Add(clearance_box, flag=wx.EXPAND|wx.ALL, border=10)
        
        # 基板エッジクリアランス設定
        board_edge_box = wx.BoxSizer(wx.HORIZONTAL)
        board_edge_label = wx.StaticText(panel, label="基板エッジクリアランス (mm):", size=(200, -1))
        board_edge_box.Add(board_edge_label, flag=wx.RIGHT|wx.ALIGN_CENTER_VERTICAL, border=8)
        self.board_edge_ctrl = wx.TextCtrl(panel, value=str(self.board_edge_clearance))
        board_edge_box.Add(self.board_edge_ctrl, proportion=1)
        panel_sizer.Add(board_edge_box, flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.BOTTOM, border=10)
        
        # ゾーンエッジクリアランス設定
        zone_box = wx.BoxSizer(wx.HORIZONTAL)
        zone_label = wx.StaticText(panel, label="ゾーンエッジクリアランス (mm):", size=(200, -1))
        zone_box.Add(zone_label, flag=wx.RIGHT|wx.ALIGN_CENTER_VERTICAL, border=8)
        self.zone_ctrl = wx.TextCtrl(panel, value=str(self.zone_clearance))
        zone_box.Add(self.zone_ctrl, proportion=1)
        panel_sizer.Add(zone_box, flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.BOTTOM, border=10)
        
        # オプション（グループ化）
        options_box = wx.StaticBox(panel, label="チェックオプション")
        options_sizer = wx.StaticBoxSizer(options_box, wx.VERTICAL)
        
        self.check_components = wx.CheckBox(panel, label="部品との衝突チェック")
        self.check_components.SetValue(True)
        options_sizer.Add(self.check_components, flag=wx.EXPAND|wx.ALL, border=5)
        
        self.check_nets = wx.CheckBox(panel, label="異なるネットとの衝突チェック")
        self.check_nets.SetValue(True)
        options_sizer.Add(self.check_nets, flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.BOTTOM, border=5)
        
        self.check_board_edge = wx.CheckBox(panel, label="基板エッジとの衝突チェック")
        self.check_board_edge.SetValue(True)
        options_sizer.Add(self.check_board_edge, flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.BOTTOM, border=5)
        
        self.check_zones = wx.CheckBox(panel, label="ゾーンエッジとの衝突チェック")
        self.check_zones.SetValue(True)
        options_sizer.Add(self.check_zones, flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.BOTTOM, border=5)
        
        # 新規オプション：基板外VIA削除
        self.check_outside_board = wx.CheckBox(panel, label="基板外のVIAを削除")
        self.check_outside_board.SetValue(True)
        options_sizer.Add(self.check_outside_board, flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.BOTTOM, border=5)
        
        panel_sizer.Add(options_sizer, flag=wx.EXPAND|wx.ALL, border=10)
        
        # パネルの設定
        panel.SetSizer(panel_sizer)
        panel.Fit()
        
        # スクロールパネルのサイズを設定
        scroll_panel.SetSizer(vbox)
        vbox.Add(panel, flag=wx.EXPAND)
        scroll_panel.FitInside()  # スクロールバーを正しく表示するために必要
        
        # ボタン（メインパネルに追加）
        button_box = wx.BoxSizer(wx.HORIZONTAL)
        cancel_button = wx.Button(main_panel, wx.ID_CANCEL, "キャンセル")
        ok_button = wx.Button(main_panel, wx.ID_OK, "OK")
        button_box.Add(cancel_button)
        button_box.Add(ok_button, flag=wx.LEFT, border=10)
        
        # メインパネルにスクロールパネルとボタンを追加
        main_sizer.Add(scroll_panel, proportion=1, flag=wx.EXPAND|wx.ALL, border=10)
        main_sizer.Add(button_box, flag=wx.ALIGN_RIGHT|wx.ALL, border=20)
        
        main_panel.SetSizer(main_sizer)
        
        # イベントバインディング
        ok_button.Bind(wx.EVT_BUTTON, self.on_ok)
        
    def on_ok(self, event):
        try:
            self.clearance = float(self.clearance_ctrl.GetValue())
            self.board_edge_clearance = float(self.board_edge_ctrl.GetValue())
            self.zone_clearance = float(self.zone_ctrl.GetValue())
            
            if self.clearance < 0 or self.board_edge_clearance < 0 or self.zone_clearance < 0:
                wx.MessageBox("クリアランスは正の値を入力してください", "エラー", wx.OK | wx.ICON_ERROR)
                return
            
            event.Skip()  # ダイアログを閉じる
        except ValueError:
            wx.MessageBox("有効な数値を入力してください", "エラー", wx.OK | wx.ICON_ERROR)

class ViaCleaner(pcbnew.ActionPlugin):
    def defaults(self):
        self.name = "VIA クリーナー"
        self.category = "編集"
        self.description = "選択したVIAから衝突や不適切なクリアランスのものを削除します"
        self.show_toolbar_button = True
        self.icon_file_name = os.path.join(os.path.dirname(__file__), 'via_cleaner.png')
    
    def Run(self):
        board = pcbnew.GetBoard()
        
        # 選択されたアイテムを取得
        selected_items = []
        for item in board.GetTracks():
            if item.IsSelected() and item.Type() == pcbnew.PCB_VIA_T:
                selected_items.append(item)
        
        if not selected_items:
            wx.MessageBox("VIAが選択されていません。VIAまたはVIAを含むグループを選択してください。", "情報", wx.OK | wx.ICON_INFORMATION)
            return
        
        # ダイアログを表示
        dialog = ViaCleanerDialog(None)
        if dialog.ShowModal() != wx.ID_OK:
            dialog.Destroy()
            return
        
        # ダイアログから設定を取得
        min_clearance = pcbnew.FromMM(dialog.clearance)
        board_edge_clearance = pcbnew.FromMM(dialog.board_edge_clearance)
        zone_clearance = pcbnew.FromMM(dialog.zone_clearance)
        check_components = dialog.check_components.GetValue()
        check_nets = dialog.check_nets.GetValue()
        check_board_edge = dialog.check_board_edge.GetValue()
        check_zones = dialog.check_zones.GetValue()
        check_outside_board = dialog.check_outside_board.GetValue()
        dialog.Destroy()
        
        # 処理開始時間
        start_time = time.time()
        
        # 基板の境界線を取得（基板エッジクリアランスチェック用）
        board_outlines = []
        if check_board_edge:
            for drawing in board.GetDrawings():
                if drawing.GetClass() == "PCB_SHAPE" and drawing.GetLayer() == pcbnew.Edge_Cuts:
                    board_outlines.append(drawing)
        
        # 基板の外接矩形（バウンディングボックス）を取得
        board_bbox = None
        outside_board_method = 0  # 基板外VIA削除の方法: 0=無効, 1=矩形, 2=個別線分
        
        if check_outside_board:
            try:
                # 方法1: バウンディングボックスを直接取得（KiCad 9系）
                if hasattr(board, "GetBoardEdgesBoundingBox"):
                    board_bbox = board.GetBoardEdgesBoundingBox()
                    if board_bbox and board_bbox.GetWidth() > 0 and board_bbox.GetHeight() > 0:
                        outside_board_method = 1
                    else:
                        board_bbox = None
                
                # 方法2: 基板コンテナからバウンディングボックスを計算（KiCad 6系互換）
                if outside_board_method == 0 and hasattr(board, "ComputeBoundingBox"):
                    board_bbox = board.ComputeBoundingBox(True)
                    if board_bbox and board_bbox.GetWidth() > 0 and board_bbox.GetHeight() > 0:
                        outside_board_method = 1
                    else:
                        board_bbox = None
                
                # 方法3: 基板エッジラインから個別に判定（最終手段）
                if outside_board_method == 0 and len(board_outlines) > 0:
                    # 基板アウトラインが存在する場合は個別線分チェックを使用
                    outside_board_method = 2
                
                # 基板外VIA削除機能の状態をユーザーに通知
                if outside_board_method == 0:
                    wx.MessageBox("基板のアウトラインが検出できませんでした。基板外VIA削除機能は無効になります。", 
                                "警告", wx.OK | wx.ICON_WARNING)
            
            except Exception as e:
                wx.MessageBox(f"基板アウトライン取得中にエラーが発生しました: {str(e)}\n基板外VIA削除機能は無効になります。", 
                            "警告", wx.OK | wx.ICON_WARNING)
                outside_board_method = 0
        
        # ゾーンを取得（ゾーンエッジクリアランスチェック用）
        zones = []
        if check_zones:
            for zone in board.Zones():
                zones.append(zone)
        
        # キャッシュの準備（高速化のため）
        footprints = []
        if check_components:
            footprints = list(board.GetFootprints())
        
        tracks_by_net = {}
        if check_nets:
            for track in board.GetTracks():
                net_code = track.GetNetCode()
                if net_code not in tracks_by_net:
                    tracks_by_net[net_code] = []
                tracks_by_net[net_code].append(track)
        
        # VIAをチェックと削除
        vias_to_remove = []
        outside_board_count = 0
        
        for via in selected_items:
            should_remove = False
            is_outside_board = False
            via_pos = via.GetPosition()
            via_net = via.GetNetCode()
            via_drill = via.GetDrill()
            via_width = via.GetWidth()
            via_radius = via_width / 2
            
            # 基板外のVIAをチェック
            if check_outside_board and not should_remove:
                if outside_board_method == 1 and board_bbox:
                    # 方法1: バウンディングボックスによるチェック
                    if not board_bbox.Contains(via_pos):
                        should_remove = True
                        is_outside_board = True
                        outside_board_count += 1
                        vias_to_remove.append(via)
                        continue
                
                elif outside_board_method == 2:
                    # 方法2: 線分ベースのチェック（点が基板の内側にあるかどうかを判定）
                    crosses = 0
                    test_point = pcbnew.wxPoint(via_pos.x, via_pos.y)
                    
                    # 境界線と交差回数をカウント
                    for outline in board_outlines:
                        if outline.GetShape() == pcbnew.SHAPE_T_SEGMENT:
                            start = outline.GetStart()
                            end = outline.GetEnd()
                            
                            # 水平線と基板アウトラインの交差をチェック
                            if ((start.y > test_point.y) != (end.y > test_point.y)) and \
                               (test_point.x < start.x + (end.x - start.x) * (test_point.y - start.y) / (end.y - start.y)):
                                crosses += 1
                    
                    # 交差回数が奇数なら内側、偶数なら外側
                    if crosses % 2 == 0:  # 外側
                        should_remove = True
                        is_outside_board = True
                        outside_board_count += 1
                        vias_to_remove.append(via)
                        continue
            
            # 部品との衝突チェック
            if check_components and not should_remove:
                for footprint in footprints:
                    fp_bbox = footprint.GetBoundingBox()
                    if fp_bbox.Contains(via_pos):
                        should_remove = True
                        break
            
            # 異なるネットとの衝突チェック
            if check_nets and not should_remove:
                # 自分のネット以外のトラックをチェック
                for net_code, tracks in tracks_by_net.items():
                    if net_code == via_net:
                        continue
                    
                    for track in tracks:
                        if track.Type() == pcbnew.PCB_TRACE_T or track.Type() == pcbnew.PCB_ARC_T:
                            clearance_needed = min_clearance + via_radius + track.GetWidth()/2
                            if track.HitTest(via_pos, clearance_needed):
                                should_remove = True
                                break
                        
                        elif track.Type() == pcbnew.PCB_VIA_T and track != via:
                            other_via_pos = track.GetPosition()
                            clearance_needed = min_clearance + via_radius + track.GetWidth()/2
                            # 距離計算
                            dx = via_pos.x - other_via_pos.x
                            dy = via_pos.y - other_via_pos.y
                            distance = math.sqrt(dx*dx + dy*dy)
                            
                            if distance < clearance_needed:
                                should_remove = True
                                break
                    
                    if should_remove:
                        break
            
            # 基板エッジとの衝突チェック
            if check_board_edge and not should_remove:
                clearance_needed = board_edge_clearance + via_radius
                
                for outline in board_outlines:
                    # 輪郭の形状に応じたチェック
                    if outline.GetShape() == pcbnew.SHAPE_T_SEGMENT:
                        # 線分の場合
                        start = outline.GetStart()
                        end = outline.GetEnd()
                        if self.distance_point_to_segment(via_pos, start, end) < clearance_needed:
                            should_remove = True
                            break
                    elif outline.GetShape() == pcbnew.SHAPE_T_ARC:
                        # 円弧の場合
                        center = outline.GetCenter()
                        radius = outline.GetRadius()
                        
                        # KiCad 9.0.2では円弧の角度取得方法が異なる
                        try:
                            # 新しいAPI (KiCad 9系)
                            start = outline.GetStart()
                            end = outline.GetEnd()
                            
                            # 始点と終点から角度を計算
                            start_vec_x = start.x - center.x
                            start_vec_y = start.y - center.y
                            end_vec_x = end.x - center.x
                            end_vec_y = end.y - center.y
                            
                            start_angle = math.degrees(math.atan2(start_vec_y, start_vec_x))
                            end_angle = math.degrees(math.atan2(end_vec_y, end_vec_x))
                            
                            # 円弧の角度差を計算
                            angle = end_angle - start_angle
                            if angle <= 0:
                                angle += 360
                            
                        except AttributeError:
                            # 旧API互換性のため (KiCad 6系以前)
                            start_angle = outline.GetArcAngleStart()
                            angle = outline.GetAngle()
                        
                        # 円弧上の最近接点との距離を計算
                        distance = self.distance_point_to_arc(via_pos, center, radius, start_angle, angle)
                        if distance < clearance_needed:
                            should_remove = True
                            break
                    elif outline.GetShape() == pcbnew.SHAPE_T_CIRCLE:
                        # 円の場合
                        center = outline.GetCenter()
                        radius = outline.GetRadius()
                        # 距離計算
                        dx = via_pos.x - center.x
                        dy = via_pos.y - center.y
                        distance = math.sqrt(dx*dx + dy*dy) - radius
                        
                        if abs(distance) < clearance_needed:
                            should_remove = True
                            break
                    elif outline.GetShape() == pcbnew.SHAPE_T_POLY:
                        # ポリゴンの場合
                        try:
                            # KiCad 9.0.2 のAPIでポリゴンとの距離を計算
                            outline_poly = outline.GetPolyShape()
                            if outline_poly.Distance(via_pos) < clearance_needed:
                                should_remove = True
                                break
                        except AttributeError:
                            # 古いバージョンの場合は処理をスキップ
                            pass
            
            # ゾーンエッジとの衝突チェック
            if check_zones and not should_remove:
                clearance_needed = zone_clearance + via_radius
                
                for zone in zones:
                    # 同じネットのゾーンはスキップ
                    if zone.GetNetCode() == via_net:
                        continue
                    
                    # ゾーン内にVIAがあるかチェック
                    try:
                        zone_poly = zone.Outline()
                        distance = zone_poly.Distance(via_pos)
                        if distance < clearance_needed:
                            should_remove = True
                            break
                    except AttributeError:
                        # 古いバージョンの場合は処理をスキップ
                        pass
            
            if should_remove and not is_outside_board:  # 基板外VIAとして既に追加済みの場合はスキップ
                vias_to_remove.append(via)
        
        # 削除を実行
        if vias_to_remove:
            for via in vias_to_remove:
                board.Remove(via)
            
            pcbnew.Refresh()
            
            # 処理時間計測
            end_time = time.time()
            execution_time = end_time - start_time
            
            # 基板外VIAの情報を含める
            if outside_board_count > 0:
                wx.MessageBox(f"{len(vias_to_remove)} 個のVIAを削除しました。\n（うち基板外VIA: {outside_board_count}個）\n処理時間: {execution_time:.2f}秒", 
                              "完了", wx.OK | wx.ICON_INFORMATION)
            else:
                wx.MessageBox(f"{len(vias_to_remove)} 個のVIAを削除しました。\n処理時間: {execution_time:.2f}秒", 
                              "完了", wx.OK | wx.ICON_INFORMATION)
        else:
            # 処理時間計測
            end_time = time.time()
            execution_time = end_time - start_time
            
            wx.MessageBox(f"削除するVIAはありませんでした。\n処理時間: {execution_time:.2f}秒", 
                          "情報", wx.OK | wx.ICON_INFORMATION)
    
    # 点と線分の距離を計算する関数
    def distance_point_to_segment(self, point, segment_start, segment_end):
        """点と線分の最短距離を計算"""
        # ベクトル計算（PCBNEWのVECTOR2Iオブジェクトに対応）
        segment_vec_x = segment_end.x - segment_start.x
        segment_vec_y = segment_end.y - segment_start.y
        
        # 線分の長さの二乗
        segment_length_squared = segment_vec_x * segment_vec_x + segment_vec_y * segment_vec_y
        
        # 線分の長さが0の場合（点の場合）
        if segment_length_squared == 0:
            dx = point.x - segment_start.x
            dy = point.y - segment_start.y
            return math.sqrt(dx*dx + dy*dy)
        
        # 点から線分への射影の比率（0〜1の範囲に制限）
        point_vec_x = point.x - segment_start.x
        point_vec_y = point.y - segment_start.y
        
        # 内積計算
        dot_product = segment_vec_x * point_vec_x + segment_vec_y * point_vec_y
        
        t = max(0, min(1, dot_product / segment_length_squared))
        
        # 線分上の最近接点
        projection_x = segment_start.x + segment_vec_x * t
        projection_y = segment_start.y + segment_vec_y * t
        
        # 点と最近接点の距離
        dx = point.x - projection_x
        dy = point.y - projection_y
        return math.sqrt(dx*dx + dy*dy)
    
    # 点と円弧の距離を計算する関数
    def distance_point_to_arc(self, point, arc_center, arc_radius, start_angle_deg, angle_deg):
        """点と円弧の最短距離を計算"""
        # 点と中心の距離
        dx = point.x - arc_center.x
        dy = point.y - arc_center.y
        center_to_point = math.sqrt(dx*dx + dy*dy)
        
        # 中心から点への角度（ラジアン）
        angle_to_point = math.atan2(dy, dx)
        # 度数法に変換
        angle_to_point_deg = math.degrees(angle_to_point)
        
        # 円弧の開始角と終了角（度数法）
        start_angle_norm = start_angle_deg % 360
        end_angle_norm = (start_angle_norm + angle_deg) % 360
        
        # 点の角度が円弧の範囲内かどうか
        is_in_range = False
        if start_angle_norm <= end_angle_norm:
            is_in_range = start_angle_norm <= angle_to_point_deg <= end_angle_norm
        else:  # 0度を跨ぐ場合
            is_in_range = angle_to_point_deg >= start_angle_norm or angle_to_point_deg <= end_angle_norm
        
        if is_in_range:
            # 円弧上の最近接点との距離
            return abs(center_to_point - arc_radius)
        else:
            # 円弧の端点との距離を計算
            start_x = arc_center.x + int(arc_radius * math.cos(math.radians(start_angle_norm)))
            start_y = arc_center.y + int(arc_radius * math.sin(math.radians(start_angle_norm)))
            
            end_x = arc_center.x + int(arc_radius * math.cos(math.radians(end_angle_norm)))
            end_y = arc_center.y + int(arc_radius * math.sin(math.radians(end_angle_norm)))
            
            # 開始点までの距離
            dx1 = point.x - start_x
            dy1 = point.y - start_y
            dist_to_start = math.sqrt(dx1*dx1 + dy1*dy1)
            
            # 終了点までの距離
            dx2 = point.x - end_x
            dy2 = point.y - end_y
            dist_to_end = math.sqrt(dx2*dx2 + dy2*dy2)
            
            return min(dist_to_start, dist_to_end)

# プラグインの登録
ViaCleaner().register()