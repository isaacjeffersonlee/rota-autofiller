import pyautogui as pa
import webbrowser
from PIL import Image
import time
import math
import numpy as np



class Autofill:
    """
    Autofill shifts according to shift_list.
    
    Attributes
    ----------
    shift_list : list
        Nested list of names and shifts e.g 
        [['Isaac Lee', ['wednesday evening', 'saturday evening', 'sunday morning']]
         ['Nithil Kennedy', ['thursday afternoon']]]

    screen_img : PngImageFile
        Pyautogui screenshot. Should be of entire monitor.

    anchor_img : PngImageFile
        Png of a point on the page to anchor to. For our
        example it is convenient to use a screenshot
        of the Staff Member cell.

    screen_region : 4d tuple[ints]
        region of the screen to be autofilled.
        i.e (top, left, width, height)
        e.g (0, 0, 1080, 1920)
        
    Methods
    -------
    """

    #                    Green:Morning  Orange:Afternoon Blue:Evening Black:Borders
    PREDEFINED_COLOURS = [(146,208,80), (248,203,173), (68,114,196), (0,0,0)]

    def __init__(self, shift_list, screen_img, anchor_img, screen_region):
        self.shift_list = shift_list
        self.screen_img = screen_img
        self.anchor_img = anchor_img
        self.screen_region = screen_region
        self.x0 = None
        self.y0 = None
    
    def check_same_colour(self, colour_one, colour_two, threshold):
        """
        Checks if two colours are the same, but perhaps
        different shades by taking the norm of the differences
        between their rgb vectors and comparing to threshold.
        Return True if below threshold and False if not.
        """
        colour_one = np.array(colour_one) 
        colour_two = np.array(colour_two) 
        if np.linalg.norm(colour_one - colour_two) < threshold:
            return True
        else:
            return False


    def get_anchor_coordinates(self):
        """
        Return the x0, y0 coordinates, which are the center
        of our anchor image.
        """
        try:
            x0, y0 = tuple(pa.locateCenterOnScreen(self.anchor_img, region=self.screen_region))
        except:
            x0, y0 = (407, 328) 

        self.x0 = x0
        self.y0 = y0
        return (x0, y0)


    def get_cells(self):
        """
        Get the y coordinate and pixel colour of all pixels in a vertical
        line, starting at the anchor image, then convert to a list of lists,
        where each sublist is a cell of a given colour.
        """
        x0, y0 = self.get_anchor_coordinates()
        max_y = self.screen_region[-1]
        pix_col_pos = [] # Initiate nested list of tuples of pixel colour and y coord
        for y in range(y0, max_y):
            pix_col = self.screen_img.getpixel((x0, y))
            if pix_col == (0,0,0):
                pix_col_pos.append((pix_col, y))
            else:
                for col in self.PREDEFINED_COLOURS:
                    # Note the second condition is so that we don't get any grays.
                    if self.check_same_colour(pix_col, col, 30) and sum(pix_col) > 100:
                        pix_col_pos.append((pix_col, y))

        black_idx = [] # Indices of all the black pixels, i.e the borders of the cells
        for i in range(len(pix_col_pos)):
            if pix_col_pos[i][0] == (0,0,0):
                black_idx.append(i)

        cells = []
        for i in range(len(black_idx) - 1):
            cell = pix_col_pos[black_idx[i]:black_idx[i+1]]
            cells.append(cell)

        # Get rid of any excess black cells
        filtered_cells = []
        for cell in cells:
            filtered_cell = []
            for pair in cell:
                if pair[0] != (0,0,0):
                    filtered_cell.append(pair)
            filtered_cells.append(filtered_cell)

        return filtered_cells

    def get_cell_midpoints(self):
        """Get the middle pixel of each cell."""
        cells = self.get_cells()
        cell_midpoints = []
        for cell in cells:
            if len(cell) > 0:
            # y coordinate of the cells midpoint
                cell_midpoint = cell[int(len(cell) / 2)]
                cell_midpoints.append(cell_midpoint)

        return cell_midpoints

    def split_into_shifts(self):
        """
        Returns list of lists of midpoints, where each
        sublist represents a shift.
        """
        cell_midpoints = self.get_cell_midpoints()
        border_idx = []
        for i in range(len(cell_midpoints)-1):
            if not self.check_same_colour(cell_midpoints[i][0], cell_midpoints[i+1][0], 30):
                border_idx.append(i)

        cells_by_shift = [cell_midpoints[:border_idx[0]+1]] # edge case
        for i in range(len(border_idx)-1):
            shift = cell_midpoints[border_idx[i]+1:border_idx[i+1]+1]
            cells_by_shift.append(shift)

        cells_by_shift.append(cell_midpoints[border_idx[-1]+1:]) # edge case

        return cells_by_shift

    def move_and_write(self, coords, text):
        """Move mouse to coords and write text."""
        pa.moveTo((coords[0], coords[1] + 1))
        pa.doubleClick()
        pa.write(text)
        pa.hotkey('esc')


    def autofill_shifts(self):
        """
        Worker function that combines the other methods
        to autofill shift_list.
        """
        shift_to_int_map = {
            'monday morning': 0, 'monday afternoon': 1, 'monday evening': 2,
            'tuesday morning': 3, 'tuesday afternoon': 4, 'tuesday evening': 5,
            'wednesday morning': 6, 'wednesday afternoon': 7, 'wednesday evening': 8,
            'thursday morning': 9, 'thursday afternoon': 10, 'thursday evening': 11,
            'friday morning': 12, 'friday afternoon': 13, 'friday evening': 14,
            'saturday morning': 15, 'saturday afternoon': 16, 'saturday evening': 17,
            'sunday morning': 18, 'sunday afternoon': 19, 'sunday evening': 20,
                            }
        shifts = self.split_into_shifts()
        person_counter = 0
        for person in self.shift_list:
            name = person[0]
            shift_sublist = person[1]
            for shift in shift_sublist:
                print(f"Filling in {shift} for {name}")
                shift_num = shift_to_int_map[shift]
                shift_cells = shifts[shift_num]
                cell = shift_cells[person_counter % len(shift_cells)] # ith person gets the ith cell
                midpoint_y = cell[1]
                coords = (self.x0, midpoint_y)
                self.move_and_write(coords=coords, text=name)

            person_counter += 1

        
def get_live_pixels():
    """Helper function."""
    time.sleep(2)
    screen_region = (0,0,1080,1920)
    screen_img = pa.screenshot(region=screen_region)
    while True:
        x,y = pa.position()
        print("")
        print(x,y)
        print(screen_img.getpixel((x,y)))

def autofill_rota(shift_list):
    screen_region = (0,0,1080,1920)
    screen_img = pa.screenshot(region=screen_region)
    anchor_img = 'staff_member.png'

    af = Autofill(shift_list=shift_list,
            screen_img=screen_img,
            anchor_img=anchor_img,
            screen_region=screen_region)

    af.autofill_shifts()


if __name__ == '__main__':
    print("Run rota_bot.py instead!")
    # get_live_pixels()
