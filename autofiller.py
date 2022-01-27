import pyautogui as pa
import webbrowser
from PIL import Image
import time
import math
import numpy as np


class Autofill:
    """
    Autofill object provides various methods to autofill a rota using
    pyautogui.

    Attributes
    ----------
    screen_region : tuple[int]
        Region of the screen to be autofilled.
        i.e (top, left, width, height)
        e.g (0, 0, 1080, 1920)

    colours : list[tuple[int]]
        List of pixel colours to filter. I.e the cell colours of the rota.
        E.g colours = [(146,208,80), (248,203,173), (68,114,196), (0,0,0)]
    """
    def __init__(self, screen_region, colours):
        self.screen_region = screen_region
        self.screen_img = pa.screenshot(region=screen_region)
        self.colours = colours
        self.y0 = None
        self.x0 = None
        self.cell_width = None
        self.cell_height = None
        self.horizontal_cell_centres = None
        self.shifts = None

    def check_same_colour(self, colour_one, colour_two, threshold=35):
        """
        Checks if two colours are the same, but perhaps
        different shades by taking the norm of the differences
        between their rgb vectors and comparing to threshold.

        Parameters
        ----------
        colour_one : tuple[int]
            (R, G, B) tuple of the first colour.

        colour_two : tuple[int]
            (R, G, B) tuple of the second colour.

        threshold : int
            The maximum distance that two colours can be 
            to be considered the same shade.
            A good value is around 35.
        
        Returns
        -------
        is_same_colour : bool
            True if same colour and False if not.
        """
        colour_one = np.array(colour_one) 
        colour_two = np.array(colour_two) 
        if np.linalg.norm(colour_one - colour_two) < threshold:
            return True
        else:
            return False


    def get_pixel_line(self, start, end, orientation, img=None):
        """
        Get all pixel colour and positions in a horizontal 
        or vertical line.
        
        Parameters
        ----------
        start : tuple[int]
            Coords of the start of the line.

        end : tuple[int]
            Coords of the end of the line.

        orientation : str 
            Orientation of the line.
            Either 'vertical' or 'horizontal'.

        img : PIL.PngImagePlugin.PngImageFile
            Image to get the pixels from.

        Returns
        -------
        pix_line : list[tuple[tuple[int], int]]
            List of tuples of pixel colour and position.
            If the line is vertical, position is the y coords
            and if horizontal, the x coords.
        """
        # If no img argument is given, we use the whole screen
        # screenshot attribute
        if not img: img = self.screen_img

        x0, y0 = start
        pix_line = []
        if orientation == 'vertical':
            for y in range(start[1], end[1]+1):
                pix_col = img.getpixel((x0, y))
                pix_line.append((pix_col, y))
        elif orientation == 'horizontal':
            for x in range(start[0], end[0]+1):
                pix_col = img.getpixel((x, y0))
                pix_line.append((pix_col, x))
        else:
            raise Exception(f"Orientation: {orientation} "\
            "is not a valid orientation! Use either vertical"\
            "or horizontal.")

        return pix_line

    def filter_pixel_line(self, pix_line):
        """
        Remove unwanted colours from the pixel line and separate
        into sub-lists, based on colour, i.e separating into separate cells.
        
        Parameters
        ----------
        pix_line : list[tuple[tuple[int], int]]
            List of tuples of pixel colour and coordinate.
        
        Returns
        -------
        filtered_cells : list[list[tuple[tuple[int], int]]]
            List of tuples of pixel colour and coordinate, grouped
            by cell colour.
        """
        filtered_pix_line = []
        # Filter out unwanted colours
        for pix in pix_line:
            pix_col, pix_coord = pix
            if pix_col == (0,0,0):
                filtered_pix_line.append(pix)
            else:
                for col in self.colours:
                    # Note the second condition is so that we don't get any grays.
                    if self.check_same_colour(pix_col, col) and sum(pix_col) > 100:
                        filtered_pix_line.append(pix)
        # If no cells detected, we return empty list
        if not filtered_pix_line: return []

        # Get all the indices of the black pixels, i.e the cell borders
        black_idx = []
        for i in range(len(filtered_pix_line)):
            if filtered_pix_line[i][0] == (0,0,0):
                black_idx.append(i)

        # If no blacks detected, return empty list
        if not black_idx: return []

        cells = []
        # First we check if we're starting in a cell
        if black_idx[0] != 0: cells.append(filtered_pix_line[0:black_idx[0]]) 
        for i in range(len(black_idx) - 1):
            cell = filtered_pix_line[black_idx[i]:black_idx[i+1]]
            cells.append(cell)

        # Get rid of any excess black cells
        filtered_cells = []
        for cell in cells:
            filtered_cell = []
            for pix in cell:
                if pix[0] != (0,0,0): # Not black
                    filtered_cell.append(pix)
            # Don't add if empty
            if filtered_cell: filtered_cells.append(filtered_cell)

        return filtered_cells


    def get_cell_centres(self, filtered_cells):
        """
        Convert the list of lists of (pix colour, pix coord) into 
        a list of (cell center pix colour, cell center y coord).
        
        Parameters
        ----------
        filtered_cells : list[list[tuple[tuple[int], int]]]
            List of tuples of pixel colour and coordinate, grouped
            by cell colour.

        Returns
        -------
        cell_centres : list[list[tuple[tuple[int], int]]]
            List of lists of (cell_centre_colour, cell_centre_y_coord) 
        """
        # Since the next cell has y coordinate += 1 of the previous
        # cell, it follows that we can just use the number of elements
        # in each cell list, i.e the number of pixels in that cell
        # to get the height of the cell.
        return [cell[int(len(cell) / 2)] for cell in filtered_cells]


    def split_into_shifts(self, cell_centres):
        """
        Get a list of list of tuples, where each sub list corresponds to a 
        a shift, and each tuple corresponds to each cell in the shift.
        In each tuple we have (cell colour, cell centre y coord).

        Parameters
        ----------
        cell_centres : list[list[tuple[tuple[int], int]]]
            List of lists of (cell_centre_colour, cell_centre_coord) 

        Returns
        -------
        cells_by_shift : list[list[tuple[tuple[int], int]]]
            List of shifts. Each shift contains tuples representing each cell
            in the shift.
        """
        border_idx = []
        for i in range(len(cell_centres)-1): 
            if not self.check_same_colour(cell_centres[i][0], cell_centres[i+1][0], 30):
                border_idx.append(i)

        cells_by_shift = [cell_centres[:border_idx[0]+1]] # edge case
        for i in range(len(border_idx)-1):
            shift = cell_centres[border_idx[i]+1:border_idx[i+1]+1]
            cells_by_shift.append(shift)

        cells_by_shift.append(cell_centres[border_idx[-1]+1:]) # edge case

        return cells_by_shift


    def get_shifts(self):
        """
        Worker function to get vertical pixel line, 
        filter it for the correct colours and call split_into_shifts 
        to parse it into shifts.

        Parameters
        ----------
        None
        
        Returns
        -------
        shifts : list[list[tuple[tuple[int], int]]]
            List of shifts. Each shift contains tuples representing each cell
            in the shift.
        """
        screen_height = self.screen_region[3]
        vertical_pix_line = self.get_pixel_line(end=(self.x0, screen_height-1),
                start=((self.x0, self.y0)),
                orientation='vertical')
        filtered_vertical_cells = self.filter_pixel_line(vertical_pix_line)
        cell_centres = self.get_cell_centres(filtered_vertical_cells)
        cells_by_shift = self.split_into_shifts(cell_centres)
        
        print(f"Shifts detected: {len(cells_by_shift)}, 21 expected.")
        return cells_by_shift

    def zoom_out(self, zooms=2):
        """
        Zoom out using ctrl-scroll wheel emulation
        with pyautogui. After zooming, pages up and 
        scrolls left.

        Parameters
        ----------
        zooms : int
            Number of mouse scrolls to emulate, equivalent
            to the number of zooms.

        Returns
        -------
        None
        """
        pa.moveTo((500,500)) # Focus the mouse
        pa.keyDown('ctrl')
        for i in range(zooms):
            pa.scroll(-1)
        pa.keyUp('ctrl')
        pa.press('pageup')
        pa.hscroll(-50)


    def calibrate_start_and_get_shifts(self):
        """
        Screenshots screen then moves down until correct colours
        are found, then checks if the correct number of shifts
        are found. If not, zooms out and recursively calls itself
        until 21 shifts are found.
        Populates the x0, y0 and cell_width and cell_height attributes.
        
        Parameters
        ----------
        None

        Returns
        -------
        None
        """
        screen_height = self.screen_region[3]
        screen_width = self.screen_region[2]
        screen_top = self.screen_region[1]
        screen_left = self.screen_region[0]
        filtered_horizontal_cells = []
        pa.click((1000,800)) # Move mouse focus
        print("Taking screenshot...")
        time.sleep(1)
        self.screen_img = pa.screenshot(region=self.screen_region) # Retake screenshot
        # Move down the screen until we detect the correct self.colours
        # Should always find in the top half, so we only go to third height
        for y in range(screen_top+200, int(screen_height/3), 10): # - constants to get negate title bars
            horizontal_pix_line = self.get_pixel_line(end=(screen_width-1, y),
                    start=((screen_left+1, y)),
                    orientation='horizontal')
            filtered_horizontal_cells = self.filter_pixel_line(horizontal_pix_line)
            if filtered_horizontal_cells:
                y0 = y + 2 # Move down a bit to avoid hitting border
                print(f"Found {len(filtered_horizontal_cells)} cells on line y={y}")
                break
            print(f"Looking for cells on the line y={y}...")
        
        if filtered_horizontal_cells: # Not-empty
            cell_width = len(filtered_horizontal_cells[-1])
            self.horizontal_cell_centres = self.get_cell_centres(filtered_horizontal_cells)
            self.cell_width = cell_width
            print(f"Updated cell_width: {self.cell_width}")
            self.y0 = y0
            print(f"Updated y0: {self.y0}")
            self.x0 = self.horizontal_cell_centres[-1][1] # Far right cell centre
            print(f"Updated x0: {self.x0}")

        else:
            pa.moveTo((500,500)) # Re-focus the mouse
            pa.hscroll(-50) # Horizontal scroll left
            time.sleep(0.5)
            pa.press('pageup')
            time.sleep(0.5)
            # Recursively call again
            print("Recursively calling calibrate_start_and_get_shifts again...")
            self.calibrate_start_and_get_shifts()

        shifts = self.get_shifts() 
        if len(shifts) == 21:
            print("All cells should now be on screen...")
            print("Updating shifts...")
            self.shifts = shifts
            # Difference between cell centre y coords for first two cells in first shift
            self.cell_height = shifts[-1][1][1] - shifts[-1][0][1]
            print("Updating cell height:", self.cell_height)

        else:
            print("Zooming out...")
            self.zoom_out()
            # Recursively call again
            self.calibrate_start_and_get_shifts()


    def move_and_write(self, coords, text):
        """
        Wrapper function which combines pyautogui doubleClick
        and write into a singe function.
        Move to coords and write text.
        
        Parameters
        ----------
        coords : tuple[int]
            (x,y) coordinates to move to.
        
        Returns
        -------
        None
        """
        pa.doubleClick((coords[0], coords[1]))
        pa.write(text)
        pa.hotkey('esc')

    def check_occupied(self, cell_centre):
        """
        Check if a cell is already occupied.
        
        Parameters
        ----------
        cell_centre : tuple[int]
            Centre (x,y) coordinate of the cell to check.
        
        Returns
        -------
        is_occupied : bool
            True if cell is already occupied, False if not.
        """
        # Get the cell region.
        cell_left = cell_centre[0] - int(self.cell_width/2)
        cell_top = cell_centre[1] - int(self.cell_height/2)
        cell_region = (cell_left, cell_top, self.cell_width, self.cell_height)
        # Screenshot only the cell_region.
        # A lot faster than screenshotting the whole screen.
        cell_img = pa.screenshot(region=cell_region)
        for y in range(1, self.cell_height-1):
            left = (1, y)
            right = (self.cell_width-1, y)
            pix_line = self.get_pixel_line(start=left, end=right, img=cell_img, orientation='horizontal')
            for i in range(len(pix_line)-1):
                # If we detect different colours, i.e someones name
                # In an empty cell we expect all the same colour
                if pix_line[i][0] != pix_line[i+1][0]:
                    return True

        return False

    def autofill_shifts(self, rota_url, shift_list):
        """
        Worker function that combines the other methods
        to autofill shift_list.

        Parameters
        ----------
        rota_url : str
            The url of the excel file to open in the browser.

        shift_list : list[list[str, list[str]]]
            Nested list of names and shifts e.g 
            [['Isaac Lee', ['wednesday evening', 'saturday evening', 'sunday morning']]
             ['Nithil Kennedy', ['thursday afternoon']]]

        Returns
        -------
        None
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
        print("Opening rota...")
        webbrowser.open(url=rota_url, new=0, autoraise=True)
        # print("Sleeping 1 sec...")
        time.sleep(3)
        # print("Finished sleeping!")
        print("Calibrating...")
        self.calibrate_start_and_get_shifts()
        displacement = 0
        failed_shifts = []
        for person in shift_list:
            name = person[0]
            shift_sublist = person[1]
            for shift in shift_sublist:
                # print("")
                # print(name)
                # print(shift)
                # print("-" * 40)
                shift_num = shift_to_int_map[shift]
                shift_cells = self.shifts[shift_num]
                # print(f"Number of cells in shift: {len(shift_cells)}")

                for i in range(len(shift_cells)):
                    # print(f"Trying cell: {i}")
                    cell = shift_cells[i] # ith person gets the ith cell
                    y = cell[1]
                    if not self.check_occupied((self.x0, y)):
                        print(f"Filling in {shift} for {name} at {(self.x0, y)}...")
                        self.move_and_write(coords=(self.x0, y), text=name)
                        break # move onto next shift
                    else: # Occupied
                        # print(f"Cell {i} for {shift} already occupied!")
                        if i == len(shift_cells)-1: # Last cell occupied
                            print(f"All cells for {shift} occupied!")
                            print("Failed to autofill shift.")
                            failed_shifts.append([name, shift])

            displacement += 1

        print("Finished autofilling.")
        print(f"No. shifts failed: {len(failed_shifts)}")
        print("Failed shifts:")
        for failed_shift in failed_shifts: print(failed_shift)


if __name__ == '__main__':
    screen_region = (0,0,1080,1920)
    shift_list = [['Isaac Lee', ['wednesday evening', 'saturday evening', 'sunday morning']],
                  ['Osaruese Egharevba', ['tuesday evening', 'wednesday afternoon', 'thursday evening', 'saturday morning', 'sunday morning']],
                  ['Nithil Kennedy', ['thursday evening']],
                  ['Rebekah Lindo', ['sunday afternoon']]]
    colours = [(146,208,80), (248,203,173), (68,114,196), (0,0,0)]
    af = Autofill(screen_region=screen_region, colours=colours)
    # af.check_occupied((349,1267))


