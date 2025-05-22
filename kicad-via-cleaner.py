#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import wx
import pcbnew
import time
import math
import json
from collections import defaultdict
from typing import List, Tuple, Dict, Set

class SpatialIndex:
    """空間インデックスによる高速近隣検索"""
    def __init__(self, grid_size=1000000):  # 1mm単位のグリッド
        self.grid_size = grid_size
        self.grid = defaultdict(list)
    
    def add_item(self, x, y, item):
        grid_x = x // self.grid_size
        grid_y = y // self.grid_size
        self.grid[(grid_x, grid_y)].append((x, y, item))
    
    def get_nearby_items(self, x, y, radius):
        grid_x = x // self.grid_size
        grid_y = y // self.grid_size
        grid_radius = (radius // self.grid_size) + 1
        
        nearby_items = []
        for gx in range(grid_x - grid_radius, grid_x + grid_radius + 1):
            for gy in range(grid_y - grid_radius, grid_y + grid_radius + 1):
                nearby_items.extend(self.grid.get((gx, gy), []))
        
        return nearby_items

class ViaCleanerDialog(wx.Dialog):
    def __init__(self, parent):
        wx.Dialog.__init__(self, parent, title="VIA クリーナー（高速化版）", size=(380, 340))
        
        # デフォルト設定
        self.default_settings = {
            'clearance': 0.2,
            'board_edge_clearance': 0.3,
            'zone_clearance': 0.2,
            'check_components': True,
            'check_nets': True,
            'check_board_edge': True,
            'check_zones': True,
            'check_outside_board': True
        }
        
        # 設定ファイルのパス
        self.settings_file = os.path.join(os.path.dirname(__file__), 'via_cleaner_settings.json')
        
        # 設定を読み込み
        self.load_settings()
        
        # メインパネル
        main_panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # ===== 数値設定部分 =====
        values_box = wx.StaticBox(main_panel, label="クリアランス設定 (mm)")
        values_sizer = wx.StaticBoxSizer(values_box, wx.VERTICAL)
        
        # グリッドレイアウトで数値設定を整理
        grid_sizer = wx.FlexGridSizer(3, 2, 5, 10)
        grid_sizer.AddGrowableCol(1, 1)
        
        # 最小クリアランス
        clearance_label = wx.StaticText(main_panel, label="最小クリアランス:")
        self.clearance_ctrl = wx.TextCtrl(main_panel, value=str(self.clearance), size=(80, -1))
        grid_sizer.Add(clearance_label, flag=wx.ALIGN_CENTER_VERTICAL)
        grid_sizer.Add(self.clearance_ctrl, flag=wx.EXPAND)
        
        # 基板エッジクリアランス
        board_edge_label = wx.StaticText(main_panel, label="基板エッジ:")
        self.board_edge_ctrl = wx.TextCtrl(main_panel, value=str(self.board_edge_clearance), size=(80, -1))
        grid_sizer.Add(board_edge_label, flag=wx.ALIGN_CENTER_VERTICAL)
        grid_sizer.Add(self.board_edge_ctrl, flag=wx.EXPAND)
        
        # ゾーンエッジクリアランス
        zone_label = wx.StaticText(main_panel, label="ゾーンエッジ:")
        self.zone_ctrl = wx.TextCtrl(main_panel, value=str(self.zone_clearance), size=(80, -1))
        grid_sizer.Add(zone_label, flag=wx.ALIGN_CENTER_VERTICAL)
        grid_sizer.Add(self.zone_ctrl, flag=wx.EXPAND)
        
        values_sizer.Add(grid_sizer, flag=wx.EXPAND|wx.ALL, border=10)
        
        # ===== チェックオプション部分 =====
        options_box = wx.StaticBox(main_panel, label="チェックオプション")
        options_sizer = wx.StaticBoxSizer(options_box, wx.VERTICAL)
        
        # チェックボックスを2列に配置
        checkbox_grid = wx.FlexGridSizer(3, 2, 5, 10)
        checkbox_grid.AddGrowableCol(0, 1)
        checkbox_grid.AddGrowableCol(1, 1)
        
        self.check_components = wx.CheckBox(main_panel, label="部品との衝突")
        self.check_components.SetValue(self.check_components_value)
        checkbox_grid.Add(self.check_components, flag=wx.EXPAND)
        
        self.check_nets = wx.CheckBox(main_panel, label="異なるネット")
        self.check_nets.SetValue(self.check_nets_value)
        checkbox_grid.Add(self.check_nets, flag=wx.EXPAND)
        
        self.check_board_edge = wx.CheckBox(main_panel, label="基板エッジ")
        self.check_board_edge.SetValue(self.check_board_edge_value)
        checkbox_grid.Add(self.check_board_edge, flag=wx.EXPAND)
        
        self.check_zones = wx.CheckBox(main_panel, label="ゾーンエッジ")
        self.check_zones.SetValue(self.check_zones_value)
        checkbox_grid.Add(self.check_zones, flag=wx.EXPAND)
        
        self.check_outside_board = wx.CheckBox(main_panel, label="基板外VIA削除")
        self.check_outside_board.SetValue(self.check_outside_board_value)
        checkbox_grid.Add(self.check_outside_board, flag=wx.EXPAND)
        
        # 空のスペースを追加してバランスを取る
        checkbox_grid.Add(wx.StaticText(main_panel, label=""), flag=wx.EXPAND)
        
        options_sizer.Add(checkbox_grid, flag=wx.EXPAND|wx.ALL, border=10)
        
        # ===== ボタン部分 =====
        button_box = wx.BoxSizer(wx.HORIZONTAL)
        
        reset_button = wx.Button(main_panel, label="デフォルト", size=(80, -1))
        button_box.Add(reset_button)
        
        # スペーサーを追加
        button_box.AddStretchSpacer()
        
        cancel_button = wx.Button(main_panel, wx.ID_CANCEL, "キャンセル", size=(80, -1))
        ok_button = wx.Button(main_panel, wx.ID_OK, "OK", size=(80, -1))
        
        button_box.Add(cancel_button, flag=wx.RIGHT, border=5)
        button_box.Add(ok_button)
        
        # メイン配置
        main_sizer.Add(values_sizer, flag=wx.EXPAND|wx.ALL, border=10)
        main_sizer.Add(options_sizer, flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.BOTTOM, border=10)
        main_sizer.Add(button_box, flag=wx.EXPAND|wx.ALL, border=15)
        
        main_panel.SetSizer(main_sizer)
        
        # イベントバインディング 
        ok_button.Bind(wx.EVT_BUTTON, self.on_ok)
        reset_button.Bind(wx.EVT_BUTTON, self.on_reset)
        
        # ダイアログを中央に配置
        self.Center()
    
    def load_settings(self):
        """設定を読み込み"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                
                # 設定値を適用（存在しない場合はデフォルト値を使用）
                self.clearance = settings.get('clearance', self.default_settings['clearance'])
                self.board_edge_clearance = settings.get('board_edge_clearance', self.default_settings['board_edge_clearance'])
                self.zone_clearance = settings.get('zone_clearance', self.default_settings['zone_clearance'])
                self.check_components_value = settings.get('check_components', self.default_settings['check_components'])
                self.check_nets_value = settings.get('check_nets', self.default_settings['check_nets'])
                self.check_board_edge_value = settings.get('check_board_edge', self.default_settings['check_board_edge'])
                self.check_zones_value = settings.get('check_zones', self.default_settings['check_zones'])
                self.check_outside_board_value = settings.get('check_outside_board', self.default_settings['check_outside_board'])
            else:
                # 設定ファイルが存在しない場合はデフォルト値を使用
                self.reset_to_defaults()
        except Exception as e:
            # 設定ファイルの読み込みに失敗した場合はデフォルト値を使用
            wx.MessageBox(f"設定ファイルの読み込みに失敗しました。デフォルト値を使用します。\nエラー: {str(e)}", 
                         "警告", wx.OK | wx.ICON_WARNING)
            self.reset_to_defaults()
    
    def save_settings(self):
        """設定を保存"""
        try:
            settings = {
                'clearance': self.clearance,
                'board_edge_clearance': self.board_edge_clearance,
                'zone_clearance': self.zone_clearance,
                'check_components': self.check_components.GetValue(),
                'check_nets': self.check_nets.GetValue(),
                'check_board_edge': self.check_board_edge.GetValue(),
                'check_zones': self.check_zones.GetValue(),
                'check_outside_board': self.check_outside_board.GetValue()
            }
            
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=2, ensure_ascii=False)
                
        except Exception as e:
            wx.MessageBox(f"設定ファイルの保存に失敗しました。\nエラー: {str(e)}", 
                         "エラー", wx.OK | wx.ICON_ERROR)
    
    def reset_to_defaults(self):
        """デフォルト値に戻す"""
        self.clearance = self.default_settings['clearance']
        self.board_edge_clearance = self.default_settings['board_edge_clearance']
        self.zone_clearance = self.default_settings['zone_clearance']
        self.check_components_value = self.default_settings['check_components']
        self.check_nets_value = self.default_settings['check_nets']
        self.check_board_edge_value = self.default_settings['check_board_edge']
        self.check_zones_value = self.default_settings['check_zones']
        self.check_outside_board_value = self.default_settings['check_outside_board']
    
    def on_reset(self, event):
        """デフォルトに戻すボタンのイベント"""
        result = wx.MessageBox("設定をデフォルト値に戻しますか？", "確認", wx.YES_NO | wx.ICON_QUESTION)
        if result == wx.YES:
            self.reset_to_defaults()
            # UIを更新
            self.clearance_ctrl.SetValue(str(self.clearance))
            self.board_edge_ctrl.SetValue(str(self.board_edge_clearance))
            self.zone_ctrl.SetValue(str(self.zone_clearance))
            self.check_components.SetValue(self.check_components_value)
            self.check_nets.SetValue(self.check_nets_value)
            self.check_board_edge.SetValue(self.check_board_edge_value)
            self.check_zones.SetValue(self.check_zones_value)
            self.check_outside_board.SetValue(self.check_outside_board_value)
            
            wx.MessageBox("設定をデフォルト値に戻しました", "完了", wx.OK | wx.ICON_INFORMATION)
        
    def on_ok(self, event):
        try:
            self.clearance = float(self.clearance_ctrl.GetValue())
            self.board_edge_clearance = float(self.board_edge_ctrl.GetValue())
            self.zone_clearance = float(self.zone_ctrl.GetValue())
            
            if self.clearance < 0 or self.board_edge_clearance < 0 or self.zone_clearance < 0:
                wx.MessageBox("クリアランスは正の値を入力してください", "エラー", wx.OK | wx.ICON_ERROR)
                return
            
            # 設定を保存
            self.save_settings()
            
            event.Skip()  # ダイアログを閉じる
        except ValueError:
            wx.MessageBox("有効な数値を入力してください", "エラー", wx.OK | wx.ICON_ERROR)

class OptimizedViaCleaner(pcbnew.ActionPlugin):
    def defaults(self):
        self.name = "VIA クリーナー（高速化版）"
        self.category = "編集"
        self.description = "選択したVIAから衝突や不適切なクリアランスのものを高速で削除します"
        self.show_toolbar_button = True
        self.icon_file_name = os.path.join(os.path.dirname(__file__), 'via_cleaner.png')
    
    def Run(self):
        board = pcbnew.GetBoard()
        
        # 選択されたVIAを取得
        selected_vias = []
        for item in board.GetTracks():
            if item.IsSelected() and item.Type() == pcbnew.PCB_VIA_T:
                selected_vias.append(item)
        
        if not selected_vias:
            wx.MessageBox("VIAが選択されていません。VIAまたはVIAを含むグループを選択してください。", "情報", wx.OK | wx.ICON_INFORMATION)
            return
        
        # ダイアログを表示
        dialog = ViaCleanerDialog(None)
        if dialog.ShowModal() != wx.ID_OK:
            dialog.Destroy()
            return
        
        # 設定取得
        min_clearance = pcbnew.FromMM(dialog.clearance)
        board_edge_clearance = pcbnew.FromMM(dialog.board_edge_clearance)
        zone_clearance = pcbnew.FromMM(dialog.zone_clearance)
        check_components = dialog.check_components.GetValue()
        check_nets = dialog.check_nets.GetValue()
        check_board_edge = dialog.check_board_edge.GetValue()
        check_zones = dialog.check_zones.GetValue()
        check_outside_board = dialog.check_outside_board.GetValue()
        dialog.Destroy()
        
        start_time = time.time()
        
        # 高速化のための前処理
        spatial_cache = self._build_spatial_cache(board, selected_vias, min_clearance, check_components, check_nets)
        board_info = self._get_board_info(board, check_board_edge, check_outside_board)
        zone_info = self._get_zone_info(board, check_zones)
        
        # VIAをチェック
        vias_to_remove = []
        outside_board_count = 0
        component_collision_count = 0
        net_collision_count = 0
        board_edge_collision_count = 0
        zone_collision_count = 0
        
        for via in selected_vias:
            removal_reason = self._check_via_fast(via, spatial_cache, board_info, zone_info, 
                                                min_clearance, board_edge_clearance, zone_clearance,
                                                check_components, check_nets, check_board_edge, 
                                                check_zones, check_outside_board)
            
            if removal_reason:
                vias_to_remove.append(via)
                if removal_reason == "outside_board":
                    outside_board_count += 1
                elif removal_reason == "component_collision":
                    component_collision_count += 1
                elif removal_reason == "net_collision":
                    net_collision_count += 1
                elif removal_reason == "board_edge_collision":
                    board_edge_collision_count += 1
                elif removal_reason == "zone_collision":
                    zone_collision_count += 1
        
        # 削除実行
        if vias_to_remove:
            for via in vias_to_remove:
                board.Remove(via)
            
            pcbnew.Refresh()
            
            end_time = time.time()
            execution_time = end_time - start_time
            
            # 詳細な削除理由を表示
            details = []
            if outside_board_count > 0:
                details.append(f"基板外VIA: {outside_board_count}個")
            if component_collision_count > 0:
                details.append(f"部品衝突: {component_collision_count}個")
            if net_collision_count > 0:
                details.append(f"ネット衝突: {net_collision_count}個")
            if board_edge_collision_count > 0:
                details.append(f"基板エッジ衝突: {board_edge_collision_count}個")
            if zone_collision_count > 0:
                details.append(f"ゾーン衝突: {zone_collision_count}個")
            
            detail_text = "\n".join(details) if details else ""
            
            message = f"{len(vias_to_remove)} 個のVIAを削除しました。\n"
            if detail_text:
                message += f"\n削除理由の詳細:\n{detail_text}\n"
            message += f"\n処理時間: {execution_time:.2f}秒"
            
            wx.MessageBox(message, "完了", wx.OK | wx.ICON_INFORMATION)
        else:
            end_time = time.time()
            execution_time = end_time - start_time
            wx.MessageBox(f"削除するVIAはありませんでした。\n処理時間: {execution_time:.2f}秒", 
                          "情報", wx.OK | wx.ICON_INFORMATION)
    
    def _build_spatial_cache(self, board, selected_vias, min_clearance, check_components, check_nets):
        """空間インデックスとキャッシュを構築"""
        cache = {}
        
        # 部品の空間インデックス（シンプル版）
        if check_components:
            footprints_list = list(board.GetFootprints())
            cache['footprints_list'] = footprints_list
        
        # トラックの空間インデックス（ネット別）
        if check_nets:
            tracks_by_net = defaultdict(lambda: SpatialIndex())
            for track in board.GetTracks():
                if track.Type() in [pcbnew.PCB_TRACE_T, pcbnew.PCB_ARC_T, pcbnew.PCB_VIA_T]:
                    net_code = track.GetNetCode()
                    if track.Type() == pcbnew.PCB_VIA_T:
                        pos = track.GetPosition()
                        tracks_by_net[net_code].add_item(pos.x, pos.y, track)
                    else:
                        # トラックの中点を使用
                        start = track.GetStart()
                        end = track.GetEnd()
                        center_x = (start.x + end.x) // 2
                        center_y = (start.y + end.y) // 2
                        tracks_by_net[net_code].add_item(center_x, center_y, track)
            cache['tracks_by_net'] = tracks_by_net
        
        return cache
    
    def _get_board_info(self, board, check_board_edge, check_outside_board):
        """基板情報を取得"""
        if not (check_board_edge or check_outside_board):
            return None
        
        board_outlines = []
        board_bbox = None
        
        # 基板アウトライン取得
        for drawing in board.GetDrawings():
            if drawing.GetClass() == "PCB_SHAPE" and drawing.GetLayer() == pcbnew.Edge_Cuts:
                board_outlines.append(drawing)
        
        # バウンディングボックス取得
        if check_outside_board:
            try:
                if hasattr(board, "GetBoardEdgesBoundingBox"):
                    board_bbox = board.GetBoardEdgesBoundingBox()
                elif hasattr(board, "ComputeBoundingBox"):
                    board_bbox = board.ComputeBoundingBox(True)
            except:
                board_bbox = None
        
        return {
            'outlines': board_outlines,
            'bbox': board_bbox
        }
    
    def _get_zone_info(self, board, check_zones):
        """ゾーン情報を取得"""
        if not check_zones:
            return None
        
        zones = []
        for zone in board.Zones():
            zones.append(zone)
        
        return {'zones': zones}
    
    def _check_via_fast(self, via, spatial_cache, board_info, zone_info, 
                       min_clearance, board_edge_clearance, zone_clearance,
                       check_components, check_nets, check_board_edge, 
                       check_zones, check_outside_board):
        """高速化されたVIAチェック"""
        via_pos = via.GetPosition()
        via_net = via.GetNetCode()
        via_radius = via.GetWidth() // 2
        
        # 基板外チェック（最も高速）
        if check_outside_board and board_info and board_info['bbox']:
            if not board_info['bbox'].Contains(via_pos):
                return "outside_board"
        
        # 部品との衝突チェック（シンプル版）
        if check_components and 'footprints_list' in spatial_cache:
            # 全部品をチェック（シンプルで確実）
            for footprint in spatial_cache['footprints_list']:
                bbox = footprint.GetBoundingBox()
                if bbox.Contains(via_pos):
                    return "component_collision"
        
        # 異なるネットとの衝突チェック（空間インデックス使用）
        if check_nets and 'tracks_by_net' in spatial_cache:
            tracks_by_net = spatial_cache['tracks_by_net']
            search_radius = min_clearance + via_radius * 2  # 余裕を持った検索半径
            
            for net_code, track_index in tracks_by_net.items():
                if net_code == via_net:
                    continue
                
                nearby_tracks = track_index.get_nearby_items(via_pos.x, via_pos.y, search_radius)
                
                for track_x, track_y, track in nearby_tracks:
                    if self._check_track_collision_fast(via, track, min_clearance):
                        return "net_collision"
        
        # 基板エッジとの衝突チェック
        if check_board_edge and board_info and board_info['outlines']:
            clearance_needed = board_edge_clearance + via_radius
            
            for outline in board_info['outlines']:
                if self._distance_to_outline_fast(via_pos, outline) < clearance_needed:
                    return "board_edge_collision"
        
        # ゾーンとの衝突チェック
        if check_zones and zone_info:
            clearance_needed = zone_clearance + via_radius
            
            for zone in zone_info['zones']:
                if zone.GetNetCode() == via_net:
                    continue
                
                try:
                    zone_poly = zone.Outline()
                    if zone_poly.Distance(via_pos) < clearance_needed:
                        return "zone_collision"
                except AttributeError:
                    pass
        
        return None  # 削除不要
    
    def _check_track_collision_fast(self, via, track, min_clearance):
        """高速トラック衝突チェック"""
        via_pos = via.GetPosition()
        via_radius = via.GetWidth() // 2
        
        if track.Type() == pcbnew.PCB_VIA_T:
            if track == via:
                return False
            
            other_pos = track.GetPosition()
            other_radius = track.GetWidth() // 2
            clearance_needed = min_clearance + via_radius + other_radius
            
            # 距離の二乗で比較（sqrt計算を回避）
            dx = via_pos.x - other_pos.x
            dy = via_pos.y - other_pos.y
            distance_squared = dx * dx + dy * dy
            clearance_needed_squared = clearance_needed * clearance_needed
            
            return distance_squared < clearance_needed_squared
        
        elif track.Type() in [pcbnew.PCB_TRACE_T, pcbnew.PCB_ARC_T]:
            clearance_needed = min_clearance + via_radius + track.GetWidth() // 2
            return track.HitTest(via_pos, clearance_needed)
        
        return False
    
    def _distance_to_outline_fast(self, point, outline):
        """高速アウトライン距離計算"""
        if outline.GetShape() == pcbnew.SHAPE_T_SEGMENT:
            return self._distance_point_to_segment_fast(point, outline.GetStart(), outline.GetEnd())
        elif outline.GetShape() == pcbnew.SHAPE_T_CIRCLE:
            center = outline.GetCenter()
            radius = outline.GetRadius()
            dx = point.x - center.x
            dy = point.y - center.y
            center_distance = math.sqrt(dx*dx + dy*dy)
            return abs(center_distance - radius)
        else:
            # 他の形状は元の処理を使用
            return self.distance_point_to_segment(point, outline.GetStart(), outline.GetEnd())
    
    def _distance_point_to_segment_fast(self, point, segment_start, segment_end):
        """高速化された点と線分の距離計算"""
        segment_vec_x = segment_end.x - segment_start.x
        segment_vec_y = segment_end.y - segment_start.y
        
        segment_length_squared = segment_vec_x * segment_vec_x + segment_vec_y * segment_vec_y
        
        if segment_length_squared == 0:
            dx = point.x - segment_start.x
            dy = point.y - segment_start.y
            return math.sqrt(dx*dx + dy*dy)
        
        point_vec_x = point.x - segment_start.x
        point_vec_y = point.y - segment_start.y
        
        dot_product = segment_vec_x * point_vec_x + segment_vec_y * point_vec_y
        t = max(0, min(1, dot_product / segment_length_squared))
        
        projection_x = segment_start.x + segment_vec_x * t
        projection_y = segment_start.y + segment_vec_y * t
        
        dx = point.x - projection_x
        dy = point.y - projection_y
        return math.sqrt(dx*dx + dy*dy)
    
    # 元のヘルパーメソッドも保持（互換性のため）
    def distance_point_to_segment(self, point, segment_start, segment_end):
        return self._distance_point_to_segment_fast(point, segment_start, segment_end)
    
    def distance_point_to_arc(self, point, arc_center, arc_radius, start_angle_deg, angle_deg):
        """元の円弧距離計算メソッド"""
        dx = point.x - arc_center.x
        dy = point.y - arc_center.y
        center_to_point = math.sqrt(dx*dx + dy*dy)
        
        angle_to_point = math.atan2(dy, dx)
        angle_to_point_deg = math.degrees(angle_to_point)
        
        start_angle_norm = start_angle_deg % 360
        end_angle_norm = (start_angle_norm + angle_deg) % 360
        
        is_in_range = False
        if start_angle_norm <= end_angle_norm:
            is_in_range = start_angle_norm <= angle_to_point_deg <= end_angle_norm
        else:
            is_in_range = angle_to_point_deg >= start_angle_norm or angle_to_point_deg <= end_angle_norm
        
        if is_in_range:
            return abs(center_to_point - arc_radius)
        else:
            start_x = arc_center.x + int(arc_radius * math.cos(math.radians(start_angle_norm)))
            start_y = arc_center.y + int(arc_radius * math.sin(math.radians(start_angle_norm)))
            
            end_x = arc_center.x + int(arc_radius * math.cos(math.radians(end_angle_norm)))
            end_y = arc_center.y + int(arc_radius * math.sin(math.radians(end_angle_norm)))
            
            dx1 = point.x - start_x
            dy1 = point.y - start_y
            dist_to_start = math.sqrt(dx1*dx1 + dy1*dy1)
            
            dx2 = point.x - end_x
            dy2 = point.y - end_y
            dist_to_end = math.sqrt(dx2*dx2 + dy2*dy2)
            
            return min(dist_to_start, dist_to_end)

# プラグインの登録
OptimizedViaCleaner().register()