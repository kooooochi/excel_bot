"""サマリーシートを追加するプロセッサー"""

from datetime import datetime
from openpyxl.workbook import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from excel_processor.base_processor import BaseSheetProcessor
from datetime import datetime


class GenerateMazeProcessor(BaseSheetProcessor):
    """
    迷路を生成し、以下のシートを生成
    - Maze : 迷路
    - Distance : Start（S）からの距離
    - Path : Start（S）からGoal（G)までの最短距離

    設定例:
        height: 10
        width: 10
    """
    def process(self, workbook: Workbook, file_path: str) -> Workbook:
        height = self.config.get('height', 10)
        width = self.config.get('width', 10)

        maze, start, goal = self._run_with_timer(
            "generate_maze",
            generate_maze,
            width=width,
            height=height
        )

        visit, path = self._run_with_timer(
            "solver",
            solver,
            maze=maze,
            start=start,
            goal=goal
        )

        self._run_with_timer(
            "output_maze_result",
            output_maze_result,
            workbook=workbook,
            maze=maze,
            start=start,
            goal=goal,
            visit=visit
        )

        return workbook

    def _run_with_timer(self, process_name, function, *args, **kwargs):
        start_time = datetime.now()
        result = function(*args, **kwargs)
        end_time = datetime.now()
        elapsed_time = end_time - start_time
        print(f"[{self.__class__.__name__}][{process_name}] 処理時間: {elapsed_time}（{elapsed_time.total_seconds():.3f} 秒）")
        return result


import random
import sys
from collections import deque
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

sys.setrecursionlimit(10**7)

def generate_maze(width, height):
    """
    width  : 迷路の横幅（奇数を推奨）
    height : 迷路の高さ（奇数を推奨）
    return : 迷路 2D 配列（壁=1, 道=0）、start座標、goal座標
    """

    maze = [[1 for _ in range(width)] for _ in range(height)]

    # 開始位置（ランダムな奇数座標）
    start_x = random.randrange(1, width, 2)
    start_y = random.randrange(1, height, 2)
    maze[start_y][start_x] = 0

    # 移動方向
    directions = [(0, 2), (0, -2), (2, 0), (-2, 0)]

    def carve(x, y):
        random.shuffle(directions)
        for dx, dy in directions:
            nx, ny = x + dx, y + dy
            if 0 < nx < width-1 and 0 < ny < height-1:
                if maze[ny][nx] == 1:
                    maze[ny - dy // 2][nx - dx // 2] = 0
                    maze[ny][nx] = 0
                    carve(nx, ny)

    carve(start_x, start_y)

    # --------------------------------------
    # Start と Goal の決定
    # --------------------------------------

    # Start = 左上の最初の通路
    S = None
    for y in range(height):
        for x in range(width):
            if maze[y][x] == 0:
                S = (x, y)
                break
        if S:
            break

    # Goal = 右下の最後の通路
    G = None
    for y in range(height - 1, -1, -1):
        for x in range(width - 1, -1, -1):
            if maze[y][x] == 0:
                G = (x, y)
                break
        if G:
            break

    return maze, S, G


def print_maze(maze, start, goal):
    sx, sy = start
    gx, gy = goal

    for y, row in enumerate(maze):
        line = ""
        for x, cell in enumerate(row):
            if (x, y) == (sx, sy):
                line += "S"
            elif (x, y) == (gx, gy):
                line += "G"
            else:
                line += "#" if cell else " "
        print(line)




class State:
    def __init__(self, xy, cost):
        self._xy = xy
        self._cost = cost
    
    def getPosition(self):
        return self._xy
    
    def getCost(self):
        return self._cost
    
    def getNextState(self, dx_dy):
        return State(
            (self._xy[0] + dx_dy[0], self._xy[1] + dx_dy[1]),
            self._cost + 1
        )
        
    
    
class Visit:
    def __init__(self, maze, start, goal):
        self._start = start
        self._goal = goal
        self._height = len(maze)
        self._width = len(maze[0])
        self._visit = [[-1 for _ in range(self._width)] for _ in range(self._height)]
        self._path = [[(-1, -1) for _ in range(self._width)] for _ in range(self._height)]
        self._maze = maze
    
    def getCost(self, xy):
        return self._visit[xy[1]][xy[0]]
    
    def setCost(self, state):
        xy = state.getPosition()
        cost = state.getCost()
        if not self.canMove(xy):
            raise ValueError(f"Invalid position {xy}. Out of bounds or wall.")
        self._visit[xy[1]][xy[0]] = cost
    
    def setPath(self, now_state, next_state):
        now_xy = now_state.getPosition()
        next_xy = next_state.getPosition()
        self._path[next_xy[1]][next_xy[0]] = now_xy
    
    def update(self, now_state, next_state):
        self.setCost(next_state)
        self.setPath(now_state, next_state)
    
    def getStartToGoalPath(self):
        path = []
        goal = tuple(self._goal)
        start = tuple(self._start)

        # 目標が未到達
        if self._path[goal[1]][goal[0]] == (-1, -1) and goal != start:
            return path

        now = goal
        while now != (-1, -1):
            path.append(now)
            if now == start:
                break
            now = self._path[now[1]][now[0]]

        return list(reversed(path))
        
    def canMove(self, xy):
        return 0<=xy[1]<self._height and 0<=xy[0]<self._width and self._maze[xy[1]][xy[0]] != 1
        
    def printVisit(self):
        for value in self._visit:
            print(value)

    def tryMove(self, now_state, next_state):
        """
        now_state から next_state への遷移が有効で未訪問なら更新する
        戻り値: True=更新した, False=スキップ
        """
        next_xy = next_state.getPosition()
        if (not self.canMove(next_xy)) or self.getCost(next_xy) >= 0:
            return False
        self.update(now_state, next_state)
        return True

    def getVisitMap(self):
        """訪問コストの2次元配列を返す（内部参照をそのまま返すので編集しないこと）"""
        return self._visit
            
    def getVisitMap(self):
        """訪問コストの2次元配列を返す（内部参照をそのまま返すので編集しないこと）"""
        return self._visit
            


def solver(maze, start, goal):
    visit = Visit(maze, start, goal)
    directions = [(1, 0), (-1, 0), (0, 1), (0, -1)]
    
    queue = deque()
    start_state = State(start, 0)
    visit.setCost(start_state)
    queue.append(start_state)
    while queue:
        now_state = queue.popleft()

        for dx_dy in directions:
            next_state = now_state.getNextState(dx_dy)
            if visit.tryMove(now_state, next_state):
                queue.append(next_state)
    
    path = visit.getStartToGoalPath()
    return visit, path


def output_maze_result(workbook, maze, start, goal, visit):
    """
    迷路、距離マップ、最短経路をそれぞれ別シートに出力する
    """
    generate_sheet_names = {
        "Maze": "Maze",
        "Distance": "Distance",
        "Path": "Path"
    }
    
    for sheet_name in generate_sheet_names.values():
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]
    
    wall_fill = PatternFill(fill_type="solid", fgColor="404040")
    start_fill = PatternFill(fill_type="solid", fgColor="4CAF50")
    goal_fill = PatternFill(fill_type="solid", fgColor="F44336")
    path_fill = PatternFill(fill_type="solid", fgColor="FFD54F")
    text_center = Alignment(horizontal="center", vertical="center")
    bold_font = Font(bold=True)
    neutral_fill = PatternFill(fill_type="solid", fgColor="FFFFFF")
    num_font = Font(color="0F172A")
    cell_size = 3  # おおよそ正方形に見える幅・高さ（単位: Excel の列幅/行高さ単位）

    # Maze シート
    maze_ws = workbook.create_sheet(generate_sheet_names["Maze"])
    maze_ws.sheet_view.showGridLines = False
    for y, row in enumerate(maze):
        for x, cell in enumerate(row):
            c = maze_ws.cell(row=y + 1, column=x + 1)
            if (x, y) == start:
                c.value = "S"
                c.fill = start_fill
                c.font = bold_font
            elif (x, y) == goal:
                c.value = "G"
                c.fill = goal_fill
                c.font = bold_font
            else:
                c.value = ""
                c.fill = wall_fill if cell else neutral_fill
            c.alignment = text_center
    for col in range(1, len(maze[0]) + 1):
        maze_ws.column_dimensions[get_column_letter(col)].width = cell_size
    for row_idx in range(1, len(maze) + 1):
        maze_ws.row_dimensions[row_idx].height = cell_size * 5  # 行高さは幅より大きめ係数

    # Distance シート
    dist_ws = workbook.create_sheet(generate_sheet_names["Distance"])
    dist_ws.sheet_view.showGridLines = False
    visit_map = visit.getVisitMap()
    max_cost = max(max(row) for row in visit_map if row) or 0
    for y, row in enumerate(visit_map):
        for x, cost in enumerate(row):
            c = dist_ws.cell(row=y + 1, column=x + 1, value=cost)
            c.alignment = text_center
            c.font = num_font
            if cost >= 0 and max_cost > 0:
                # 簡易ヒートマップ: コストに応じて薄い青から濃い青へ
                intensity = int(255 - (cost / max_cost) * 120)
                hex_part = f"{intensity:02X}"
                c.fill = PatternFill(fill_type="solid", fgColor=f"BB{hex_part}FF")
            else:
                c.fill = wall_fill
    for col in range(1, len(maze[0]) + 1):
        dist_ws.column_dimensions[get_column_letter(col)].width = cell_size
    for row_idx in range(1, len(maze) + 1):
        dist_ws.row_dimensions[row_idx].height = cell_size * 5

    # Path シート
    path_ws = workbook.create_sheet(generate_sheet_names["Distance"])
    path_ws.sheet_view.showGridLines = False
    path = visit.getStartToGoalPath()
    step_lookup = {xy: idx for idx, xy in enumerate(path)}
    for y, row in enumerate(maze):
        for x, cell in enumerate(row):
            coord = (x, y)
            c = path_ws.cell(row=y + 1, column=x + 1)
            if coord == start:
                c.value = "S"
                c.fill = start_fill
                c.font = bold_font
            elif coord == goal:
                c.value = "G"
                c.fill = goal_fill
                c.font = bold_font
            elif coord in step_lookup:
                c.value = step_lookup[coord]
                c.fill = path_fill
                c.font = num_font
            elif cell == 1:
                c.value = ""
                c.fill = wall_fill
            else:
                c.value = ""
                c.fill = neutral_fill
            c.alignment = text_center
    for col in range(1, len(maze[0]) + 1):
        path_ws.column_dimensions[get_column_letter(col)].width = cell_size
    for row_idx in range(1, len(maze) + 1):
        path_ws.row_dimensions[row_idx].height = cell_size * 5