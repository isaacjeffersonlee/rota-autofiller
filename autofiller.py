import pyautogui as pa
import webbrowser
from PIL import Image
import time
import math
import numpy as np

class Autofill:
    """
    Attributes
    ----------
    shift_list : list[list[str, list[str]]]
        Nested list of names and shifts e.g 
        [['Isaac Lee', ['wednesday evening', 'saturday evening', 'sunday morning']]
         ['Nithil Kennedy', ['thursday afternoon']]]

    screen_region : tuple[int]
        region of the screen to be autofilled.
        i.e (top, left, width, height)
        e.g (0, 0, 1080, 1920)

    colours : list[tuple[int]]
        list of pixel colours to keep. I.e the cell colours
        of the rota.
        E.g colours = [(146,208,80), (248,203,173), (68,114,196), (0,0,0)]
        
    Methods
    -------
    # TODO: Add methods doc
    """
    def __init__(self, shift_list, screen_region, colours):
        self.shift_list = shift_list 
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
        colour_one : 3d tuple
            (R, G, B) tuple of the first colour.

        colour_two : 3d tuple
            (R, G, B) tuple of the second colour.

        threshold : float
            The maximum distance that two colours can be 
            to be considered the same shade.
            A good value is around 35.
        
        Returns
        -------
        name : type
            Return value description
        """
        colour_one = np.array(colour_one) 
        colour_two = np.array(colour_two) 
        if np.linalg.norm(colour_one - colour_two) < threshold:
            return True
        else:
            return False


    def get_pixel_line(self, start, end, orientation):
        """
        Get all pixel colour and positions in a horizontal 
        or vertical line.
        
        Parameters
        ----------
        start : tuple
            Coords of the start of the line.

        end : tuple
            Coords of the end of the line.

        orientation : str 
            Orientation of the line.
            Either 'vertical' or 'horizontal'.

        Returns
        -------
        pix_line : list
            List of lists of pixel colour and position.
            If the line is vertical, position is the y coords
            and if horizontal, the x coords.
        """
        x0, y0 = start
        pix_line = []
        if orientation == 'vertical':
            for y in range(start[1], end[1]+1):
                pix_col = self.screen_img.getpixel((x0, y))
                pix_line.append((pix_col, y))
        elif orientation == 'horizontal':
            for x in range(start[0], end[0]+1):
                pix_col = self.screen_img.getpixel((x, y0))
                pix_line.append((pix_col, x))
        else:
            raise Exception(f"Orientation: {orientation} "\
            "is not a valid orientation! Use either vertical"\
            "or horizontal.")

        return pix_line


    def filter_pixel_line(self, pix_line):
        """
        Remove unwanted self.colours from the pixel line and separate
        pixel_line into sub-lists, based on colour, i.e separating
        into separate cells.
        
        Parameters
        ----------
        pix_line : list
            list of tuples of pixel colour and coordinate.
        
        Returns
        -------
        filtered_cells : list
            list of lists of pixel colour and coordinate based, grouped
            by cell colour.
        """
        filtered_pix_line = []
        # Filter out unwanted self.colours
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
        a list of (cell center pix colour, cell center coord).
        
        Parameters
        ----------
        filtered_cells : list
            List of lists of pixel colour and coordinate based, grouped
            by cell colour.

        Returns
        -------
        cell_centres : list
            List of lists of (cell_centre_colour, cell_centre_coord) 
        """
        cell_centres = []
        for cell in filtered_cells:
            cell_centre = cell[int(len(cell) / 2)]
            cell_centres.append(cell_centre)

        return cell_centres


    def split_into_shifts(self, cell_centres):
        """
        Returns list of lists of midpoints, where each
        sublist represents a shift.
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
        """Returns a list of vertical cells, starting at (x0,y0)."""
        screen_height = self.screen_region[3]
        vertical_pix_line = self.get_pixel_line(end=(self.x0, screen_height-1),
                start=((self.x0, self.y0)),
                orientation='vertical')
        filtered_vertical_cells = self.filter_pixel_line(vertical_pix_line)
        cell_centres = self.get_cell_centres(filtered_vertical_cells)
        shifts = self.split_into_shifts(cell_centres)
        
        print(f"Shifts detected: {len(shifts)}, 21 expected.")
        return shifts


    def zoom_out(self, zooms=4):
        """
        Zoom out using ctrl-scroll wheel emulation
        with pyautogui.
        """
        pa.moveTo((500,500)) # Focus the mouse
        # time.sleep(0.5)
        pa.keyDown('ctrl')
        for i in range(zooms):
            pa.scroll(-1)
        # time.sleep(0.5)
        pa.keyUp('ctrl')
        # time.sleep(0.5)
        pa.press('pageup')
        # time.sleep(0.5)
        pa.hscroll(-50)
        # time.sleep(0.5)


    def calibrate_start_and_get_shifts(self):
        """
        Move down the screen until correct colours are found.
        Populate the x0, y0 and cell width attributes.
        """
        screen_height = self.screen_region[3]
        screen_width = self.screen_region[2]
        screen_top = self.screen_region[1]
        screen_left = self.screen_region[0]
        filtered_horizontal_cells = []
        # Move down the screen until we detect the correct self.colours
        # Should always find in the top half, so we only go to third height
        for y in range(screen_top+200, int(screen_height/3), 10): # - constants to get negate title bars
            horizontal_pix_line = self.get_pixel_line(end=(screen_width-1, y),
                    start=((screen_left, y)),
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
            print("Retaking screenshot...")
            self.screen_img = pa.screenshot(region=self.screen_region) # Retake screenshot
            # Recursively call again
            print("Recursively calling calibrate_start_and_get_shifts again...")
            self.calibrate_start_and_get_shifts()

        shifts = self.get_shifts() 
        if len(shifts) == 21:
            print("All cells should now be on screen...")
            print("Updating shifts...")
            self.shifts = shifts
            # Difference between cell centre y coords for first two cells in first shift
            self.cell_height =  shifts[0][1][1] - shifts[0][0][1]
            print("Updating cell height:", self.cell_height)

        else:
            print("Zooming out...")
            self.zoom_out()
            # Recursively call again
            print("Retaking screenshot...")
            self.screen_img = pa.screenshot(region=self.screen_region) # Retake screenshot
            self.calibrate_start_and_get_shifts()


    def move_and_write(self, coords, text):
        """Double click coords and write text."""
        pa.doubleClick((coords[0], coords[1]))
        pa.write(text)
        pa.hotkey('esc')

    def check_occupied(self, cell_centre):
        """Check if a cell already has text."""
        cell_left = cell_centre[0] - int(self.cell_width/2)
        cell_top = cell_centre[1] - int(self.cell_height/2)
        cell_region = (cell_left, cell_top, self.cell_width, self.cell_height)
        self.screen_img = pa.screenshot(region=cell_region)
        for y in range(self.cell_height):
            left = (1, y)
            right = (self.cell_width-1, y)
            pix_line = self.get_pixel_line(start=left, end=right, orientation='horizontal')
            for i in range(len(pix_line)-1):
                if pix_line[i][0] != pix_line[i+1][0]:
                    return True

        return False

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
        self.calibrate_start_and_get_shifts()
        displacement = 0
        failed_shifts = []
        for person in self.shift_list:
            name = person[0]
            shift_sublist = person[1]
            for shift in shift_sublist:
                print("")
                print(name)
                print(shift)
                print("-" * 40)
                shift_num = shift_to_int_map[shift]
                shift_cells = self.shifts[shift_num]
                print(f"No. cells in shift: {len(shift_cells)}")

                for i in range(len(shift_cells)):
                    print(f"Trying cell: {i}")
                    cell = shift_cells[i] # ith person gets the ith cell
                    y = cell[1]
                    if not self.check_occupied((self.x0, y)):
                        print(f"Filling in {shift} for {name} at {(self.x0, y)}...")
                        self.move_and_write(coords=(self.x0, y), text=name)
                        break # move onto next shift
                    else: # Occupied
                        print(f"Cell {i} for {shift} already occupied!")
                        if i == len(shift_cells)-1: # Last cell occupied
                            print(f"All cells for {shift} occupied!")
                            print("Failed to autofill shift")
                            failed_shifts.append([name, shift])

            displacement += 1

        print("Finished autofilling...")
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
    af = Autofill(shift_list=shift_list, screen_region=screen_region, colours=colours)
    af.autofill_shifts()
    # af.check_occupied((349,1267))


