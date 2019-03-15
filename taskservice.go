package taskscheduler

import (
	"errors"
	"time"

	ole "github.com/go-ole/go-ole"
	oleutil "github.com/go-ole/go-ole/oleutil"
)

// Task is a task found in Windows Task Scheduler 2.0
type Task struct {
	Name        string
	Path        string
	Enabled     bool
	LastRunTime time.Time
	NextRunTime time.Time
	ActionList  []ExecAction // Other actions are ignored, we are only interested in Commandline Actions
}

// ExecAction is an action defined in a scheduled Task if type IExecAction.
type ExecAction struct {
	WorkingDirectory string
	Path             string
	Arguments        string
}

// GetTasks returns a list of all scheduled Tasks in Windows Task Scheduler 2.0
func GetTasks() ([]Task, error) {
	// Initialize COM API
	if err := ole.CoInitialize(0); err != nil {
		return nil, errors.New("Could not initialize Windows COM API")
	}
	defer ole.CoUninitialize()
	// Create an ITaskService object
	unknown, err := ole.CreateInstance(ole.NewGUID("{0F87369F-A4E5-4CFC-BD3E-73E6154572DD}"), nil)
	if err != nil {
		return nil, errors.New("Could not initialize Task Scheduler 2.0")
	}
	defer unknown.Release()
	// Convert IUnknown to IDispatch to get more functions like CallMethod()
	ts, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, errors.New("Could not prepare Task Scheduler 2.0")
	}
	defer ts.Release()
	// Connect to the Task Scheduler 2.0
	if _, err := ts.CallMethod("Connect", "", "", "", ""); err != nil {
		return nil, errors.New("Could not connect to Task Scheduler 2.0")
	}
	// Get Root Directory of Task Scheduler 2.0 and get all tasks recursively
	root := oleutil.MustCallMethod(ts, "GetFolder", "\\").ToIDispatch()
	defer root.Release()
	return getTasksRecursively(root), nil
}

func getTasksRecursively(folder *ole.IDispatch) (tasks []Task) {
	// Get Tasks in subfolders first
	folderIterator := oleutil.MustCallMethod(folder, "GetFolders", int64(0)).ToIDispatch()
	for i := int32(1); i <= oleutil.MustGetProperty(folderIterator, "count").Value().(int32); i++ {
		// Get Tasks of subfolder i
		index := ole.NewVariant(ole.VT_I4, int64(i))
		subfolder := oleutil.MustGetProperty(folderIterator, "item", &index).ToIDispatch()
		subtasks := getTasksRecursively(subfolder)
		tasks = append(tasks, subtasks...)
		subfolder.Release()
	}
	folderIterator.Release()
	// Get Tasks
	taskIterator := oleutil.MustCallMethod(folder, "GetTasks", int64(0)).ToIDispatch()
	for i := int32(1); i <= oleutil.MustGetProperty(taskIterator, "count").Value().(int32); i++ {
		// Get Task i
		index := ole.NewVariant(ole.VT_I4, int64(i))
		task := oleutil.MustGetProperty(taskIterator, "item", &index).ToIDispatch()
		var t = Task{
			Name:        oleutil.MustGetProperty(task, "name").ToString(),
			Path:        oleutil.MustGetProperty(task, "path").ToString(),
			Enabled:     oleutil.MustGetProperty(task, "enabled").Value().(bool),
			LastRunTime: oleutil.MustGetProperty(task, "nextRunTime").Value().(time.Time),
			NextRunTime: oleutil.MustGetProperty(task, "lastRunTime").Value().(time.Time),
		}
		// Get more details, e.g. actions
		definition := oleutil.MustGetProperty(task, "definition").ToIDispatch()
		actions := oleutil.MustGetProperty(definition, "actions").ToIDispatch()
		count := oleutil.MustGetProperty(actions, "count").Value().(int32)
		for i := int32(1); i <= count; i++ {
			// Get Action i
			index := ole.NewVariant(ole.VT_I4, int64(i))
			action := oleutil.MustGetProperty(actions, "item", &index).ToIDispatch()
			if oleutil.MustGetProperty(action, "type").Value().(int32) != 0 { // only handle IExecAction
				action.Release()
				continue
			}
			t.ActionList = append(t.ActionList, ExecAction{
				WorkingDirectory: oleutil.MustGetProperty(action, "workingDirectory").ToString(),
				Path:             oleutil.MustGetProperty(action, "path").ToString(),
				Arguments:        oleutil.MustGetProperty(action, "arguments").ToString(),
			})
			action.Release()
		}
		tasks = append(tasks, t)
		task.Release()
	}
	taskIterator.Release()
	return
}
