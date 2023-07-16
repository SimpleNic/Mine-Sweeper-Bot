import pyautogui 
import time
import win32api, win32con
import win32gui
import webbrowser

CWIDTH = 12
CHEIGHT = 12

global_tiles=0
global_unrevealed_tiles=0
global_unrevealed_bombs=0
global_bomb_chance=0
global_row=0
global_column=0

class Tile:
    def __init__(self,square):
        self.x=square[1].left+CWIDTH
        self.y=square[1].top+CHEIGHT
        self.num = 0
        self.vir_num = 0
        self.row = square[0] // global_column
        self.column = square[0] % global_column
        self.bomb_chance = -1

    def read_around(self):
        tiles_around = []
        for r in range(-1,2):
            row = self.row+r
            if row < 0 or row > global_row-1: continue
            for c in range(-1,2):
                column = self.column+c
                if column < 0 or column > global_column-1 or (c==0 and r==0): continue
                pos = (row)*global_column + column
                tiles_around.append(global_tiles[pos])
        return tiles_around
    
    def flag(self):
        global global_unrevealed_bombs,global_bomb_chance
        global_unrevealed_bombs -= 1
        global_bomb_chance = global_unrevealed_bombs/global_unrevealed_tiles

        Rclick(self.x,self.y)
        self.num = 11
        tiles_around = self.read_around()
        tiles_changed = []
        for tile_around in tiles_around:
            if tile_around.num in range(1,9):
                tile_around.vir_num -= 1
                tiles_changed.append(tile_around)
        return tiles_changed

    def reveal(self):
        Lclick(self.x,self.y)
        self.update()
        tiles_changed = self.update_mass_tiles()
        return tiles_changed

    def update(self):
        global global_unrevealed_tiles,global_bomb_chance
        global_unrevealed_tiles -= 1
        if global_unrevealed_tiles != 0:
            global_bomb_chance = global_unrevealed_bombs/global_unrevealed_tiles
        
        if self.num != 0: raise Exception('How did you get this?')

        color = win32gui.GetPixel(win32gui.GetDC(win32gui.GetActiveWindow()), self.x+1,self.y)
        
        # 0                             unrevealed
        # 1 0,  38, 247     16197120
        # 2 55, 120,30      1996855
        # 3 232,52, 18      1193192
        # 4 0,  12, 119     7801856
        # 5 111,20, 4       267375
        # 6 55, 121,123     8091959
        # 7 0,  0,  0       6184542
        # 8 123,123,123     8092539
        # 9 189,189,189     12434877    revealed empty
        # 10                460551      bomb
        # 11                            flag

        match color:
            case 16197120:
                self.num=1
            case 1996855:
                self.num=2
            case 1193192:
                self.num=3
            case 7801856:
                self.num=4
            case 267375:
                self.num=5
            case 8091959:
                self.num=6
            case 6184542:
                self.num=7
            case 8092539:
                self.num=8
            case 12434877:
                self.num=9
            case 0:
                self.num=10
            case _:
                if global_bomb_chance != 0: 
                    win32api.SetCursorPos((global_tiles[0].x,global_tiles[0].y))
                    time.sleep(.1)
                    win32api.SetCursorPos((self.x,self.y))
                    exc = f'Color: {color}, Row: {self.row}, Column: {self.column}'
                    raise Exception(exc)
        self.vir_num = self.num

    def update_mass_tiles(self):  # Can speed up maybe by updating whole board
        tiles_changed = []
        if self.num in range(1,9): 
            tiles_changed.append(self)
            tiles_around = self.read_around()
            for tile_around in tiles_around:
                if tile_around.num in range(1,9):
                    tiles_changed.append(tile_around)
                elif tile_around.num == 11:
                    self.vir_num -= 1
            return tiles_changed
        elif self.num != 9: return []

        tiles_around = self.read_around()
        for tile_around in tiles_around:
            if tile_around.num != 0: continue
            tile_around.update()
            tmp = tile_around.update_mass_tiles()
            tiles_changed.extend(tmp)

        return tiles_changed

def Lclick(x,y):
    x = (int)(x)
    y = (int)(y)
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    time.sleep(.01)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)
    time.sleep(.05)

def Rclick(x,y):
    x = (int)(x)
    y = (int)(y)
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,0,0)
    time.sleep(.01)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,0,0)
    time.sleep(.01)

def setup():
    global global_tiles,global_unrevealed_tiles,global_unrevealed_bombs,global_bomb_chance,global_row,global_column
    level = input('Enter the level of difficulty: beginner, intermediate, expert\n')
    difficulty = {'beginner':[10,9,9],'intermediate':[40,16,16],'expert':[99,16,30],'':[99,16,30]}
    if (list)(difficulty.keys()).count(level) == 0: raise Exception('You must input a valid difficulty: beginner, intermediate, expert')
    webbrowser.open(f'https://minesweeperonline.com/#{level}')

    global_unrevealed_bombs = difficulty[level][0]
    global_row = difficulty[level][1]
    global_column = difficulty[level][2]
    squares = []
    while len(squares) == 0:
        time.sleep(1)
        squares = (list)(pyautogui.locateAllOnScreen('C:\\Users\\nicol\\Documents\\Code\\images\\space.png', grayscale=True, confidence=.9))
    
    tiles = []
    for square in enumerate(squares):
        tiles.append(Tile(square))

    global_tiles = tiles
    global_unrevealed_tiles = len(global_tiles)
    global_bomb_chance = global_unrevealed_bombs/global_unrevealed_tiles

def print_board():
    for tile in global_tiles:
        print(tile.num,end=' ')
        if tile.num < 10:
            print(' ',end='')
        if tile.column == 29:
            print('\n')

def print_board_chances():
    for tile in global_tiles:
        if tile.bomb_chance != -1:
            print('',end=' ')
        print(f'{tile.bomb_chance:.1f}',end=' ')
        if tile.column == 29:
            print('\n')

def get_tiles_unrevealed(tiles):
    tiles_unrevealed = []
    for tile in tiles:
        if tile.num == 0:
            tiles_unrevealed.append(tile)
    return tiles_unrevealed

def get_tiles_numbered(tiles):
    tiles_numbered = []
    for tile in tiles:
        if tile.num in range(1,9):
            tiles_numbered.append(tile)
    return tiles_numbered

def get_tiles_sets(tiles1,tiles2):
    tiles_both = []
    tiles_unique1 = []
    tiles_unique2 = []
    for tile1 in tiles1:
        if tiles2.__contains__(tile1):
            tiles_both.append(tile1)
    for tile1 in tiles1:
        if not tiles_both.__contains__(tile1):
            tiles_unique1.append(tile1)
    for tile2 in tiles2:
        if not tiles_both.__contains__(tile2):
            tiles_unique2.append(tile2)

    return tiles_unique1,tiles_unique2,tiles_both

def calculate_simple_chances(tile,tiles_around):
    tiles_unrevealed_around = []

    for tile_around in tiles_around:
        if tile_around.num == 0:
            tiles_unrevealed_around.append(tile_around)

    if len(tiles_unrevealed_around) == 0: return -1

    bomb_chance_around = tile.vir_num/len(tiles_unrevealed_around)
    for tile_unrevealed_around in tiles_unrevealed_around:
        if tile_unrevealed_around.bomb_chance < bomb_chance_around:
            tile_unrevealed_around.bomb_chance = bomb_chance_around

    return bomb_chance_around

def calculate_set_chances(tile1,tile2,tiles_unique1,tiles_unique2,tiles_both):
    big = tile1
    small = tile2
    set_diff_size = len(tiles_unique1)
    bomb_chance_around = 1
    if tile1.vir_num < tile2.vir_num: 
        big = tile2
        small = tile1
        set_diff_size = len(tiles_unique2)
        bomb_chance_around = 0

    if (big.vir_num - small.vir_num) != set_diff_size: return -1

    for tile_both in tiles_both:
        tile_both.bomb_chance = (big.vir_num-set_diff_size)/len(tiles_both)

    return bomb_chance_around

def solve():
    tiles_changed = []
    tiles_unused = []
    tiles_changed.extend(global_tiles[(global_column-1)*(global_row-1)//2 + (global_column)//2].reveal())

    while len(tiles_changed) > 0:
        tile = tiles_changed.pop()
        is_not_used = True
        tiles_around = tile.read_around()
        bomb_chance = calculate_simple_chances(tile,tiles_around)

        for tile_around in tiles_around:
            if tile_around.num != 0: continue
            if bomb_chance == 0:
                tiles_changed.extend(tile_around.reveal())
                is_not_used = False
            elif bomb_chance == 1:
                tiles_changed.extend(tile_around.flag())
                is_not_used = False
        if is_not_used: tiles_unused.append(tile)

        while len(tiles_changed) == 0 and global_bomb_chance != 0 and len(tiles_unused) > 0:
            tile = tiles_unused.pop()
            tiles_unrevealed = get_tiles_unrevealed(tile.read_around())
            if len(tiles_unrevealed) == 0: continue
            for tile_unrevealed in tiles_unrevealed:
                tiles_number = get_tiles_numbered(tile_unrevealed.read_around())
                if len(tiles_number) < 2: continue
                for i in range(len(tiles_number)-1):
                    for j in range(1,len(tiles_number)):
                        tile1 = tiles_number[i]
                        tile2 = tiles_number[j]
                        tiles_unrevealed1 = get_tiles_unrevealed(tile1.read_around())
                        tiles_unrevealed2 = get_tiles_unrevealed(tile2.read_around())
                        tiles_unique1,tiles_unique2,tiles_both = get_tiles_sets(tiles_unrevealed1,tiles_unrevealed2)
                        bomb_chance = calculate_set_chances(tile1,tile2,tiles_unique1,tiles_unique2,tiles_both)
                        
                        if bomb_chance == 0:
                            for tile_unique1 in tiles_unique1:
                                if tile_unique1.num != 0: continue
                                tiles_changed.extend(tile_unique1.reveal())
                            for tile_unique2 in tiles_unique2:
                                tiles_changed.extend(tile_unique2.flag())

                        elif bomb_chance == 1:
                            for tile_unique1 in tiles_unique1:
                                tiles_changed.extend(tile_unique1.flag())
                            for tile_unique2 in tiles_unique2:
                                if tile_unique2.num != 0: continue
                                tiles_changed.extend(tile_unique2.reveal())

        if len(tiles_changed) == 0 and len(tiles_unused) == 0:
            tiles_unrevealed = get_tiles_unrevealed(global_tiles)
            lowest_chance = tiles_unrevealed[0]
            for tile in tiles_unrevealed:
                if tile.num != 0: continue
                if tile.bomb_chance == -1: tile.bomb_chance = global_bomb_chance
                if lowest_chance.bomb_chance > tile.bomb_chance: lowest_chance = tile
            tiles_changed.extend(tile.reveal())

        if tile.num == 10: return False
    return True


if __name__ == '__main__':
    setup()
    if solve(): print('Yay!')
    else: print('Oh no!')
    win32api.SetCursorPos((global_tiles[0].x,global_tiles[0].y))

